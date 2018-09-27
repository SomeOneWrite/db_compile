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
