import logging
import re

class ColorUtil:

    @staticmethod
    def ShadeColor(hexColor, percent):
        try:
            hexNum = int(hexColor[1:], 16)
            t = 0 if percent < 0 else 255
            p = percent * -1 if percent < 0 else percent

            R = hexNum >> 16
            G = hexNum >> 8 & 0x00FF
            B = hexNum & 0x0000FF

            # print(hexColor, percent)
            # print(R, G, B)

            Rhex = hex(round((t - R) * p) + R)[2:].zfill(2)
            Ghex = hex(round((t - G) * p) + G)[2:].zfill(2)
            BHex = hex(round((t - B) * p) + B)[2:].zfill(2)

            # print(Rhex, Ghex, BHex)

            hexString = '#' + Rhex + Ghex + BHex

            # print(hexString)
            # print()

            return hexString
        except Exception as e:
            logging.error(logging.exception("Error creating shade of color"))

    # https://stackoverflow.com/questions/20275524/how-to-check-if-a-string-is-an-rgb-hex-string
    @staticmethod
    def IsValidHexColor(hexColor):
        hexstring = re.compile(r'#[a-fA-F0-9]{3}(?:[a-fA-F0-9]{3})?$')
        return bool(hexstring.match(hexColor))
