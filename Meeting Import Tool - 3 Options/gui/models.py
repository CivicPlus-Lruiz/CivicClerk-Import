"""
Data Models
MeetingBody, Subtype, FileError
"""


class MeetingBody:
    """Represents a meeting body configuration."""
    def __init__(self, folder_path: str = "", meeting_type: str = "",
                 event_category: str = "", default_time: str = ""):
        self.folder_path = folder_path
        self.meeting_type = meeting_type
        self.event_category = event_category
        self.default_time = default_time

    def to_dict(self) -> dict:
        return {
            "folder_path": self.folder_path,
            "meeting_type": self.meeting_type,
            "event_category": self.event_category,
            "default_time": self.default_time
        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            folder_path=data.get("folder_path", ""),
            meeting_type=data.get("meeting_type", ""),
            event_category=data.get("event_category", ""),
            default_time=data.get("default_time", "")
        )


class Subtype:
    """Represents a subtype/alias configuration."""
    def __init__(self, event_category: str = "", folder_name: str = "",
                 subtype: str = "", aliases: str = "", override_time: str = "",
                 override_meeting_type: str = "", event_category_override: str = ""):
        self.event_category = event_category
        self.folder_name = folder_name
        self.subtype = subtype
        self.aliases = aliases
        self.override_time = override_time
        self.override_meeting_type = override_meeting_type
        self.event_category_override = event_category_override

    def to_dict(self) -> dict:
        return {
            "event_category": self.event_category,
            "folder_name": self.folder_name,
            "subtype": self.subtype,
            "aliases": self.aliases,
            "override_time": self.override_time,
            "override_meeting_type": self.override_meeting_type,
            "event_category_override": self.event_category_override
        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            event_category=data.get("event_category", ""),
            folder_name=data.get("folder_name", ""),
            subtype=data.get("subtype", ""),
            aliases=data.get("aliases", ""),
            override_time=data.get("override_time", ""),
            override_meeting_type=data.get("override_meeting_type", ""),
            event_category_override=data.get("event_category_override", "")
        )


class FileError:
    """Represents a file error that needs resolution."""
    def __init__(self, error_type: str, file_name: str, issue: str,
                 original_location: str, current_location: str, file_path: str = ""):
        self.error_type = error_type
        self.file_name = file_name
        self.issue = issue
        self.original_location = original_location
        self.current_location = current_location
        self.file_path = file_path
        self.selected = False
