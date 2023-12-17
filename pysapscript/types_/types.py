from enum import Enum


class NavigateAction(Enum):
    """
    Type for Window.navigate()
    """

    enter = "enter"
    back = "back"
    end = "end"
    cancel = "cancel"
    save = "save"
