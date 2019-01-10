

class ColorIndex:
    xlColorIndexAutomatic = -4105  # from enum XlColorIndex
    xlColorIndexNone = -4142  # from enum XlColorIndex


def int_to_rgb(number):
    """Given an integer, return the rgb"""
    number = int(number)
    r = number % 256
    g = (number // 256) % 256
    b = (number // (256 * 256)) % 256
    return r, g, b


def rgb_to_int(rgb):
    """Given an rgb, return an int"""
    return rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
