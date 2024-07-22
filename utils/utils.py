# njcaa url
DEFAULT_URL = 'https://www.njcaa.org'

# njcaa teams url
TEAMS_URL = 'https://www.njcaa.org/sports/bsb/teams'

"""
This method is to convert into int datatype
"""
def safe_int(value: str) -> any:
    """Attempt to convert a value to an integer, returning a default if conversion fails."""
    try:
        return int(value)
    except ValueError:
        return value


"""
This method is to convert into float datatype
"""
def safe_float(value: str) -> any:
    """Attempt to convert a value to a float, returning a default if conversion fails."""
    try:
        return float(value)
    except ValueError:
        return value