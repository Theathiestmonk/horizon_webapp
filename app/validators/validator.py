class Severity:
    value: str = "error"

class Issue:
    def __init__(self, message: str):
        self.message = message
        self.severity = Severity()

def validate_payload(payload: dict) -> list:
    """Dummy validation service"""
    return []
