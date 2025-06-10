"""Exceptions thrown"""


class WindowDidNotAppearException(Exception):
    """Main windows didn't show up - possible pop-up window"""


class AttachException(Exception):
    """Error with attaching - connection or session"""


class ActionException(Exception):
    """Error performing action - click, select ..."""

class ActionNotAllowedException(Exception):
    """Action not allowed - e.g. sending VKey to a non-editable field"""
