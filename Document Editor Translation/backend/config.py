# ================================================================
# config.py
# This module handles environment variable loading, logging setup,
# and configuration file parsing. It also builds a mapping of model names
# to their corresponding codes.
# ================================================================

import json  # For parsing JSON files
import logging  # For logging events and errors

from dotenv import load_dotenv  # To load environment variables from a .env file

# Load environment variables from the .env file, if it exists.
load_dotenv()

# ------------------------------------------------
# Setup Logging
# ------------------------------------------------
# Logging is configured to write messages to ".log".
# This file will record informational messages and errors.
if not logging.getLogger().handlers:
    logging.basicConfig(
        filename=".log",  # The log file path
        level=logging.INFO,  # Log level; change to DEBUG for more detailed output
        format="%(asctime)s - %(levelname)s - %(message)s",  # Format for each log message
    )
logging.info("Logging configured in config.py")

# ------------------------------------------------
# Load and Parse Configuration
# ------------------------------------------------
# Attempt to open and parse the JSON configuration file "data.json".
# The configuration is expected to contain a "models" key.
try:
    with open("data.json", "r", encoding="utf-8") as f:
        CONFIG = json.load(f)  # Load the JSON data into a Python dictionary
    logging.info("Configuration loaded successfully from data.json")
except Exception as e:
    # If any error occurs during loading, log the error and set default empty values.
    logging.error(f"Error loading configuration: {e}")
    CONFIG = {}
