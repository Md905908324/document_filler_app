import os

def ensure_folder(path: str):
    """Create folder if it does not exist."""
    if not os.path.exists(path):
        os.makedirs(path)
