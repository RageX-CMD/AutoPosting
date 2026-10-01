#!/usr/bin/env python3
"""Fest konfigurierte App-Variante für News-Bereich → Aktionen."""

from __future__ import annotations

import forum_poster

forum_poster.configure_target("news")

from app import main  # noqa: E402 - Ziel muss vor dem UI-Import gesetzt werden


if __name__ == "__main__":
    main()
