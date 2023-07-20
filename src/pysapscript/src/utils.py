import os
from pathlib import Path
from datetime import datetime as dt

def get_environ_var(name: str) -> str:
    return os.environ[name]
  
def get_unix_timestamp(input_datetime: dt=None) -> int:
    if input_datetime is None:
        input_datetime = dt.utcnow()

    return int(input_datetime.timestamp())

class Errors:
    def __init__(self, error_file_path: Path):
        self._error_file_path = error_file_path
    
    def write_error(self, message: str):
        with self._error_file_path.open('a') as f:
            f.write(message)