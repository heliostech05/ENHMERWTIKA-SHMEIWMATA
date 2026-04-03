#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Compatibility wrapper for the weekly package entrypoint.
"""

from weekly.timologia import main, timologia_weekly

__all__ = ["main", "timologia_weekly"]


if __name__ == "__main__":
    main()
