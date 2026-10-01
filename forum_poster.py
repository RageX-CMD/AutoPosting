"""Sichtbare Browser-Automatisierung für das deutsche Naruto-Online-Forum."""

from __future__ import annotations

import json
import os
import re
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from generate_event_docx import GenerateResult, chunk_to_forum_html


FORUM_ORIGIN = "https://forum-narutode.narutowebgame.com"
BOARD_ID = 38
BOARD_NAME = "Allgemeines"
BOARD_URL = f"{FORUM_ORIGIN}/page/show-thread-{BOARD_ID}-1.html"
CATEGORY_ID = 4
CATEGORY_NAME = "Aktionen"
STATE_FILENAME = "posting_state.json"
MAX_HTML_LENGTH = 20_000
# Das Forum lehnt Beiträge ab, wenn zwischen zwei Posts weniger als
# 30 Sekunden liegen. Fünf Sekunden Puffer vermeiden Rundungs-/Netzwerkfehler.
REPLY_DELAY_SECONDS = 35
SESSION_COOKIE_FILENAME = "forum_session.dat"

ProgressCallback = Callable[[str], None]
ManualActionCallback = Callable[[str], bool]


def configure_target(target: str) -> None:
    """Konfiguriert das feste Forumziel vor dem Import der Oberfläche."""
    global BOARD_ID, BOARD_NAME, BOARD_URL
    global CATEGORY_ID, CATEGORY_NAME, STATE_FILENAME

    if target == "news":
        BOARD_ID = 37
        BOARD_NAME = "News-Bereich"
        BOARD_URL = f"{FORUM_ORIGIN}/page/show-thread-{BOARD_ID}-1.html"
        CATEGORY_ID = 1
        CATEGORY_NAME = "Aktionen"
        STATE_FILENAME = "posting_state_news.json"
        return
    if target == "general":
        BOARD_ID = 38
        BOARD_NAME = "Allgemeines"
        BOARD_URL = f"{FORUM_ORIGIN}/page/show-thread-{BOARD_ID}-1.html"
        CATEGORY_ID = 4
        CATEGORY_NAME = "Aktionen"
        STATE_FILENAME = "posting_state.json"
        return
    raise ValueError(f"Unbekanntes Forumziel: {target}")


class ForumPostingError(RuntimeError):
    """Fehler, bei dem keine weiteren Beiträge gesendet werden dürfen."""


class LoginRequiredError(ForumPostingError):
    pass


class DuplicateThreadError(ForumPostingError):
    pass


class PostingCancelledError(ForumPostingError):
    pass


@dataclass
class PostingState:
    sheet: str
    title: str
    board_id: int
    category_id: int
    total_segments: int
    posted_segments: int = 0
    thread_id: int | None = None
    thread_url: str | None = None
    status: str = "pending"
    updated_at: str = ""

    @property
    def complete(self) -> bool:
        return self.posted_segments >= self.total_segments and self.status == "complete"


@dataclass
class PostingResult:
    state: PostingState
    browser_name: str


def persistent_edge_profile_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data)
    else:
        base = Path.home() / "AppData" / "Local"
    return base / "EventPostGenerator" / "EdgeProfile"


def persistent_chrome_profile_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data)
    else:
        base = Path.home() / "AppData" / "Local"
    return base / "EventPostGenerator" / "ChromeProfile"


def session_cookie_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data)
    else:
        base = Path.home() / "AppData" / "Local"
    return base / "EventPostGenerator" / SESSION_COOKIE_FILENAME


def protect_for_current_windows_user(data: bytes) -> bytes:
    if os.name != "nt":
        raise ForumPostingError("Windows-DPAPI ist auf diesem System nicht verfügbar.")

    import ctypes
    from ctypes import wintypes

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
        ]

    buffer = ctypes.create_string_buffer(data)
    input_blob = DataBlob(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
    )
    output_blob = DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DataBlob),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    cryptprotect_ui_forbidden = 0x1
    success = crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        "EventPostGenerator Forum Session",
        None,
        None,
        None,
        cryptprotect_ui_forbidden,
        ctypes.byref(output_blob),
    )
    if not success:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(output_blob.pbData)


def unprotect_for_current_windows_user(data: bytes) -> bytes:
    if os.name != "nt":
        raise ForumPostingError("Windows-DPAPI ist auf diesem System nicht verfügbar.")

    import ctypes
    from ctypes import wintypes

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
        ]

    buffer = ctypes.create_string_buffer(data)
    input_blob = DataBlob(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
    )
    output_blob = DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(DataBlob),
        ctypes.c_void_p,
        ctypes.POINTER(DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(DataBlob),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    cryptprotect_ui_forbidden = 0x1
    success = crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        cryptprotect_ui_forbidden,
        ctypes.byref(output_blob),
    )
    if not success:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(output_blob.pbData)


def posting_state_path(week_dir: Path) -> Path:
    return week_dir / STATE_FILENAME


def load_posting_state(week_dir: Path) -> PostingState | None:
    path = posting_state_path(week_dir)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return PostingState(**data)
    except (OSError, TypeError, ValueError, json.JSONDecodeError) as exc:
        raise ForumPostingError(f"Statusdatei ist ungültig: {path.name} ({exc})") from exc


def save_posting_state(week_dir: Path, state: PostingState) -> None:
    state.updated_at = datetime.now(timezone.utc).isoformat()
    path = posting_state_path(week_dir)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(
        json.dumps(asdict(state), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(path)


class ForumPoster:
    def __init__(
        self,
        progress: ProgressCallback | None = None,
        manual_action: ManualActionCallback | None = None,
    ) -> None:
        self.progress = progress or (lambda _message: None)
        self.manual_action = manual_action or (lambda _message: False)
        self.driver = None
        self.browser_name = ""
        self.profile_dir: Path | None = None

    def close(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def post(
        self,
        generated: GenerateResult,
        *,
        resume: bool = False,
        dry_run: bool = False,
    ) -> PostingResult:
        html_segments = [chunk_to_forum_html(chunk) for chunk in generated.chunks]
        for index, content in enumerate(html_segments, start=1):
            if len(content) > MAX_HTML_LENGTH:
                raise ForumPostingError(
                    f"Beitrag {index} überschreitet das Forum-Limit "
                    f"({len(content)} statt maximal {MAX_HTML_LENGTH} HTML-Zeichen)."
                )

        state = load_posting_state(generated.week_dir)
        if state is not None:
            self._validate_state(state, generated)
            if state.complete:
                raise DuplicateThreadError(
                    f"Diese Woche wurde bereits vollständig gepostet: {state.thread_url}"
                )
            if not resume:
                raise ForumPostingError(
                    "Es existiert ein unvollständiger Posting-Status. "
                    "Zum Fortsetzen ist eine Bestätigung erforderlich."
                )
        else:
            state = PostingState(
                sheet=generated.post.source_sheet,
                title=generated.post.title,
                board_id=BOARD_ID,
                category_id=CATEGORY_ID,
                total_segments=len(html_segments),
            )

        self._start_browser()
        self.progress(f"{self.browser_name} wurde geöffnet.")
        self._ensure_logged_in()

        if state.thread_id is None:
            self._open_board()
            duplicate_url = self._find_duplicate_thread(generated.post.title)
            if duplicate_url:
                raise DuplicateThreadError(
                    f"Im Bereich {BOARD_NAME} existiert bereits ein Thema mit diesem Titel: "
                    f"{duplicate_url}"
                )

            self._fill_thread_preview(generated.post.title, html_segments[0])
            if dry_run:
                self.progress(
                    "Dry-Run fertig: Thema ist im Browser vorbereitet und wurde nicht abgesendet."
                )
                return PostingResult(state=state, browser_name=self.browser_name)

            self.progress(f"Erstelle das neue Thema im Bereich {BOARD_NAME} …")
            thread_url = self._submit_new_thread()
            thread_id = self._thread_id_from_url(thread_url)
            state.thread_url = thread_url
            state.thread_id = thread_id
            state.posted_segments = 1
            state.status = "posting"
            save_posting_state(generated.week_dir, state)
            self.driver.get(thread_url)
            self.progress(f"Thema erstellt: {thread_url}")

        start_index = state.posted_segments
        for segment_index in range(start_index, len(html_segments)):
            number = segment_index + 1
            if state.posted_segments > 0:
                self.progress(
                    f"Warte {REPLY_DELAY_SECONDS} Sekunden vor der nächsten Antwort …"
                )
                time.sleep(REPLY_DELAY_SECONDS)
            self.progress(f"Poste Antwort {number}/{len(html_segments)} …")
            latest_url = self._submit_reply(
                state.thread_url
                or f"{FORUM_ORIGIN}/page/show-post-{state.thread_id}-1.html",
                html_segments[segment_index],
            )
            if latest_url:
                state.thread_url = latest_url
            state.posted_segments = number
            state.status = "posting"
            save_posting_state(generated.week_dir, state)

        state.status = "complete"
        save_posting_state(generated.week_dir, state)
        if state.thread_url:
            self.driver.get(state.thread_url)
        self.progress(
            f"Automatisches Posting abgeschlossen: {state.posted_segments} Beiträge."
        )
        return PostingResult(state=state, browser_name=self.browser_name)

    def _validate_state(self, state: PostingState, generated: GenerateResult) -> None:
        expected = {
            "sheet": generated.post.source_sheet,
            "title": generated.post.title,
            "board_id": BOARD_ID,
            "category_id": CATEGORY_ID,
            "total_segments": len(generated.chunks),
        }
        for field_name, expected_value in expected.items():
            if getattr(state, field_name) != expected_value:
                raise ForumPostingError(
                    f"Posting-Status passt nicht zur aktuellen Ausgabe: {field_name}."
                )
        if state.posted_segments < 0 or state.posted_segments > state.total_segments:
            raise ForumPostingError("Posting-Status enthält eine ungültige Segmentanzahl.")
        if state.posted_segments and (not state.thread_id or not state.thread_url):
            raise ForumPostingError("Posting-Status enthält keine gültige Thread-Referenz.")

    def _start_browser(self) -> None:
        if self.driver is not None:
            return
        try:
            from selenium import webdriver

            profile_dir = persistent_edge_profile_dir()
            profile_dir.mkdir(parents=True, exist_ok=True)
            edge_options = webdriver.EdgeOptions()
            edge_options.add_argument("--start-maximized")
            edge_options.add_argument("--disable-blink-features=AutomationControlled")
            edge_options.add_argument(f"--user-data-dir={profile_dir}")
            edge_options.add_argument("--profile-directory=Default")
            self.driver = webdriver.Edge(options=edge_options)
            self.browser_name = "Microsoft Edge"
            self.profile_dir = profile_dir
            self.progress(f"Dauerhaftes Edge-Profil: {profile_dir}")
            return
        except Exception as edge_error:
            edge_error_text = str(edge_error).lower()
            if (
                "user data directory is already in use" in edge_error_text
                or "profile in use" in edge_error_text
            ):
                raise ForumPostingError(
                    "Das dauerhafte Edge-Profil wird bereits verwendet. Bitte alle "
                    "Edge-Fenster des Profils „EventPostGenerator“ schließen und erneut "
                    "versuchen."
                ) from edge_error
            self.progress(f"Edge konnte nicht gestartet werden: {edge_error}")

        try:
            from selenium import webdriver

            profile_dir = persistent_chrome_profile_dir()
            profile_dir.mkdir(parents=True, exist_ok=True)
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument(f"--user-data-dir={profile_dir}")
            chrome_options.add_argument("--profile-directory=Default")
            self.driver = webdriver.Chrome(options=chrome_options)
            self.browser_name = "Google Chrome"
            self.profile_dir = profile_dir
            self.progress(f"Dauerhaftes Chrome-Profil: {profile_dir}")
        except Exception as chrome_error:
            raise ForumPostingError(
                "Weder Microsoft Edge noch Google Chrome konnte gestartet werden. "
                f"Chrome-Fehler: {chrome_error}"
            ) from chrome_error

    def _ensure_logged_in(self) -> None:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.support.ui import WebDriverWait

        self._open_board()
        if self._is_logged_in():
            self._save_session_cookies()
            self.progress("Bereits im Forum angemeldet.")
            return

        if self._restore_session_cookies():
            self.driver.refresh()
            self._wait_document_ready()
            if self._is_logged_in():
                self.progress("Gespeicherte Forum-Sitzung wurde wiederhergestellt.")
                return

        if not self.manual_action(
            f"Edge ist im Bereich „{BOARD_NAME}“ geöffnet. Die App öffnet das "
            "Loginfenster absichtlich nicht, damit Cloudflare nicht ausgelöst wird. "
            "Bitte im Browser selbst auf „Anmelden“ klicken und den kompletten "
            "Login manuell abschließen."
        ):
            raise PostingCancelledError("Login wurde abgebrochen.")

        try:
            WebDriverWait(self.driver, 120).until(lambda _driver: self._is_logged_in())
        except TimeoutException as exc:
            raise LoginRequiredError(
                "Der Forum-Login wurde nicht bestätigt. Es wurde nichts gepostet."
            ) from exc
        self._save_session_cookies()
        self.progress("Manueller Login bestätigt.")

    def _save_session_cookies(self) -> None:
        cookies = [
            cookie
            for cookie in self.driver.get_cookies()
            if "narutowebgame.com" in cookie.get("domain", "")
        ]
        if not cookies:
            self.progress("Keine Forum-Cookies zum Speichern gefunden.")
            return

        payload = json.dumps(
            {"version": 1, "cookies": cookies},
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        encrypted = protect_for_current_windows_user(payload)
        path = session_cookie_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        temporary.write_bytes(encrypted)
        temporary.replace(path)
        self.progress("Forum-Sitzung wurde verschlüsselt für diesen Windows-Benutzer gespeichert.")

    def _restore_session_cookies(self) -> bool:
        path = session_cookie_path()
        if not path.exists():
            return False
        try:
            payload = json.loads(
                unprotect_for_current_windows_user(path.read_bytes()).decode("utf-8")
            )
            if payload.get("version") != 1 or not isinstance(payload.get("cookies"), list):
                raise ValueError("Unbekanntes Cookie-Dateiformat")
        except Exception as exc:
            self.progress(f"Gespeicherte Forum-Sitzung konnte nicht gelesen werden: {exc}")
            return False

        allowed_keys = {
            "name",
            "value",
            "path",
            "domain",
            "secure",
            "httpOnly",
            "expiry",
            "sameSite",
        }
        restored = 0
        for source_cookie in payload["cookies"]:
            if not isinstance(source_cookie, dict):
                continue
            cookie = {
                key: value
                for key, value in source_cookie.items()
                if key in allowed_keys
            }
            if not cookie.get("name") or "value" not in cookie:
                continue
            try:
                self.driver.add_cookie(cookie)
                restored += 1
            except Exception:
                continue
        return restored > 0

    def _is_logged_in(self) -> bool:
        try:
            return bool(
                self.driver.execute_script(
                    "return window.guest === 0 || "
                    "!!document.querySelector('#logoutBtn, .forum_login_after');"
                )
            )
        except Exception:
            return False

    def _open_board(self) -> None:
        self.driver.get(BOARD_URL)
        self._wait_document_ready()

    def _wait_document_ready(self, timeout: float = 30) -> None:
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            try:
                if self.driver.execute_script("return document.readyState") == "complete":
                    return
            except Exception:
                pass
            time.sleep(0.2)
        raise ForumPostingError("Die Forumsseite wurde nicht vollständig geladen.")

    def _find_duplicate_thread(self, title: str) -> str | None:
        return self.driver.execute_script(
            """
            const title = arguments[0];
            const link = [...document.querySelectorAll('a')].find(
              a => a.textContent.trim() === title &&
                   a.href.includes('/page/show-post-')
            );
            return link ? link.href : null;
            """,
            title,
        )

    def _fill_thread_preview(self, title: str, content: str) -> None:
        prepared = self.driver.execute_script(
            """
            const title = arguments[0], content = arguments[1];
            const categoryName = arguments[2], categoryId = arguments[3];
            const titleInput = document.querySelector('#threadName');
            const category = document.querySelector('#category');
            const editor = document.querySelector('.wangEditor-txt');
            if (!titleInput || !category || !editor) return false;
            document.querySelector('.forum_editor_mask')?.remove();
            titleInput.value = title;
            titleInput.dispatchEvent(new Event('input', {bubbles: true}));
            category.value = categoryName;
            category.dataset.value = String(categoryId);
            if (window.jQuery) window.jQuery(category).data('value', categoryId);
            editor.innerHTML = content;
            editor.dispatchEvent(new Event('input', {bubbles: true}));
            editor.dispatchEvent(new Event('change', {bubbles: true}));
            const sendButton = document.querySelector('#sendBtn');
            if (sendButton) {
              sendButton.dataset.bOk = '1';
              if (window.jQuery) window.jQuery(sendButton).data('bOk', 1);
            }
            editor.scrollIntoView({behavior: 'smooth', block: 'center'});
            return true;
            """,
            title,
            content,
            CATEGORY_NAME,
            CATEGORY_ID,
        )
        if not prepared:
            raise ForumPostingError("Der Editor für ein neues Thema wurde nicht gefunden.")

    def _submit_new_thread(self) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait

        previous_time = self.driver.execute_script("return performance.timeOrigin")
        self.driver.find_element(By.ID, "sendBtn").click()
        try:
            WebDriverWait(self.driver, 60).until(
                lambda driver: (
                    "/page/show-post-" in driver.current_url
                    and driver.execute_script("return performance.timeOrigin")
                    != previous_time
                )
            )
        except Exception as exc:
            raise self._native_submission_error("Thema konnte nicht erstellt werden") from exc
        self._raise_if_cloudflare_blocked()
        return self.driver.current_url

    def _submit_reply(self, thread_url: str, content: str) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait

        self.driver.get(thread_url)
        self._wait_document_ready()
        self._raise_if_cloudflare_blocked()
        prepared = self.driver.execute_script(
            """
            const content = arguments[0];
            const editor = document.querySelector('.wangEditor-txt');
            const sendButton = document.querySelector('#sendBtn');
            if (!editor || !sendButton) return false;
            document.querySelector('.forum_editor_mask')?.remove();
            editor.innerHTML = content;
            editor.dispatchEvent(new Event('input', {bubbles: true}));
            editor.dispatchEvent(new Event('change', {bubbles: true}));
            sendButton.dataset.bOk = '1';
            if (window.jQuery) window.jQuery(sendButton).data('bOk', 1);
            editor.scrollIntoView({behavior: 'smooth', block: 'center'});
            return true;
            """,
            content,
        )
        if not prepared:
            raise ForumPostingError("Der Antwort-Editor wurde nicht gefunden.")

        previous_time = self.driver.execute_script("return performance.timeOrigin")
        self.driver.find_element(By.ID, "sendBtn").click()
        try:
            WebDriverWait(self.driver, 60).until(
                lambda driver: (
                    driver.execute_script("return performance.timeOrigin")
                    != previous_time
                )
            )
        except Exception as exc:
            raise self._native_submission_error("Antwort konnte nicht gepostet werden") from exc
        self._raise_if_cloudflare_blocked()
        return self.driver.current_url

    def _raise_if_cloudflare_blocked(self) -> None:
        try:
            title = self.driver.title.lower()
            body = self.driver.find_element("tag name", "body").text.lower()
        except Exception:
            return
        if (
            "attention required" in title
            or "sorry, you have been blocked" in body
            or "cloudflare ray id" in body
        ):
            raise ForumPostingError(
                "Cloudflare hat diese Browser-Sitzung blockiert. Bitte Edge vollständig "
                "schließen, einige Minuten warten und den Lauf erneut starten."
            )

    def _native_submission_error(self, prefix: str) -> ForumPostingError:
        self._raise_if_cloudflare_blocked()
        message = self.driver.execute_script(
            """
            const candidates = [
              '.forum_copyLink_tips',
              '.forum_modal_unblock_content',
              '.dlog_oas_froum_alert',
              '.handle_frame_tips'
            ];
            for (const selector of candidates) {
              const element = document.querySelector(selector);
              if (element && element.textContent.trim()) {
                return element.textContent.trim();
              }
            }
            return '';
            """
        )
        if message:
            return ForumPostingError(f"{prefix}: {message}")
        return ForumPostingError(
            f"{prefix}. Das Forum hat innerhalb von 60 Sekunden keine Bestätigung geliefert."
        )

    @staticmethod
    def _thread_id_from_url(url: str) -> int:
        match = re.search(r"/show-post-(\d+)-", url)
        if not match:
            raise ForumPostingError(f"Thread-ID konnte nicht aus der URL gelesen werden: {url}")
        return int(match.group(1))
