"""Read appraisal parameters out of the DfT TAG Data Book workbook.

Read a workbook, or use the built-in v2.03 figures without one::

    from tag_databook import TagDataBook

    book = TagDataBook.from_workbook("tag-data-book-v2-03.xlsm")
    book.vot_work_rail  # 38.215...
    TagDataBook.with_defaults().vot_work_rail  # 38.22
"""

from __future__ import annotations

from tag_databook.databook import Fallback, TagDataBook
from tag_databook.tables import CarbonValues, DiscountRates, GDPDeflator, ValueOfTime

__all__ = [
    "CarbonValues",
    "DiscountRates",
    "Fallback",
    "GDPDeflator",
    "TagDataBook",
    "ValueOfTime",
]

__version__ = "0.1.0"
