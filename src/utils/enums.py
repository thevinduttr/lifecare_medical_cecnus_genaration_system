from enum import Enum

class Final_Status(Enum):
    STARTED = "STARTED"
    PENDING = "PENDING"
    FAILED = "FAILED"
    PARTIALLY_COMPLETED = "Partially Completed"
    COMPLETED = "COMPLETED"

class Orient_Status(Enum):
    APPROVED = "APPROVED"
    REFERRAL = "REFERRAL"
    REJECTED = "REJECTED"
    NOT_APPLICABLE = "N/A"

