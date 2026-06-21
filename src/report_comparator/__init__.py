from importlib.metadata import PackageNotFoundError, version

from .comparator import compare_runs

try:
    __version__ = version("report-comparator")
except PackageNotFoundError:  # not installed (e.g. running from a source checkout)
    __version__ = "0.0.0+unknown"

__all__ = ["compare_runs", "__version__"]
