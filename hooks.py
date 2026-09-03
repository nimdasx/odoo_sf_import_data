"""Backward-compatibility shim.

The core import engine has been refactored into `tools/import_engine.py`.
This module re-exports public functions, classes, and constants so external callers
(such as client config modules using `post_init_hook`) continue to work seamlessly.
"""
from .tools.import_engine import (  # noqa: F401
    ImportLogger,
    import_bundled_data,
    run_import,
    MODULE,
    JOURNAL_TYPES,
    ASSET_METHODS,
    ASSET_PERIODS,
    ACCOUNT_TYPES,
)
