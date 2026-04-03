__all__ = ["main", "timologia_weekly"]


def __getattr__(name):
    if name == "main":
        from .timologia import main
        return main
    if name == "timologia_weekly":
        from .timologia import timologia_weekly
        return timologia_weekly
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
