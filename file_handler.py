REQUIRED_COLUMNS = [
    "Summary",
    "Issue Type",
    "Priority",
    "Status",
    "Created"
]

def validate_file(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]

    if missing:
        return False, f"Missing columns: {missing}"

    return True, "Valid file"
