"""pytest configuration — add the parent directory to sys.path so tests
can import file_parser, api_client, wbs_writer, and config modules."""

import sys
import os

# Add the main/ directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
