import re

def without_whitespace(str):
    return ''.join(str.split())

def to_float(str):
    try:
        return float(without_whitespace(str.replace(',', '.')))
    except Exception as e:
        # print('Not float value {} '.format(str))
        return None

def to_float_or_zero(str):
    try:
        return float(without_whitespace(str.replace(',', '.')))
    except Exception as e:
        # print('Not float value {} '.format(str))
        return float(0.0)

def to_float_or(str, default: float):
    try:
        return float(without_whitespace(str.replace(',', '.')))
    except Exception as e:
        # print('Not float value {} '.format(str))
        return default

def without_lines(str):
    str = re.sub(r'\n', '', str)
    str = re.sub(r'\t', '', str)
    str = re.sub(r' +', ' ', str)
    str = str.strip()
    return str
