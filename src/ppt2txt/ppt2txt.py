import argparse
import os
import sys
import re
import json
import olefile


# Reference: https://msopenspecs.azureedge.net/files/MS-PPT/%5bMS-PPT%5d.pdf
TEXT_HEADER_ATOM = "9F0F"
TEXT_BYTES_ATOM = "A80F"
ENCODING = "unicode_escape"
"""
Example:

hex_data = "9F0F04000000040000000000A80F0E0000005468697320697320612074657374"
    - 9F0F denotes the start of a TextHeaderAtom
    - the subsequent 04000000, which should be read as 00000004, is the record's length
    - 04 is the textType (see TextTypeEnum)
    - we then have 4 bits for the recVersion and 12 bits for the recInstance
    - A80F denotes the start of a TextBytesAtom
    - the subsequent 0E000000, which should be read as 0000000E is the length
      in bytes of the array of text (i.e., 14 in this case)
    - 5468697320697320612074657374 is the actual text (28 characters),
      which if decoded results in the "This is a test" string
"""


def process_args():
    parser = argparse.ArgumentParser(description='A pure python-based utility '
                                                 'to extract text '
                                                 'from binary PPT files.')
    parser.add_argument("ppt", help="path to the ppt file")
    parser.add_argument("-o", "--output_dir", help="path to the output directory", default=".")
    args = parser.parse_args()

    if not os.path.exists(args.ppt):
        print(f"File {args.ppt} does not exist")
        sys.exit(1)

    return args


def hexdump(src, length=16) -> str:
    """
    Returns a hexadecimal dump of a binary string
    (adapted from https://github.com/decalage2/oletools/blob/master/oletools/ezhexviewer.py)

    Args:
        :param src: stream source
        :param length: number of bytes per row
    """
    hex_data = ''
    for i in range(0, len(src), length):
        s = src[i : i + length]
        hex_data += ''.join(["%02X" % x for x in s])
    return hex_data


def process(path: str) -> dict:
    """
    Extract textual content from binary .ppt

    Args:
        :param path: path to the PPT
    """

    # Check if path exists
    if not os.path.exists(path):
        print(f"File {path} does not exist")
        return {}

    try:
        parsed_ppt_dict = {}

        ole = olefile.OleFileIO(path)
        meta = ole.get_metadata()
        parsed_ppt_dict["filename"] = path
        parsed_ppt_dict["slides"] = meta.slides
        parsed_ppt_dict["content"] = {}
        stream = ole.openstream('PowerPoint Document').getvalue()

        hex_data = hexdump(stream)

        matches = list(re.finditer(TEXT_HEADER_ATOM, hex_data))
        matches_spans = [match.span() for match in matches]

        text_counter = 0
        for j in range(len(matches_spans)):
            start_index = int(matches_spans[j][1])
        
            # Check if there is a TextBytesAtom
            is_text_bytes_atom = hex_data[start_index+20:start_index+24] == TEXT_BYTES_ATOM
            if not is_text_bytes_atom:
                continue

            start_index += 8
            # text_type = hex_data[start_index:start_index+2]

            start_index += 16
            # Get the bytes indicating the length of the text
            rec_len_bytes = hex_data[start_index:start_index+8]
            rec_len_bytes_couples = list(zip(rec_len_bytes[0::2], rec_len_bytes[1::2]))
            rec_len_bytes = "".join(["".join(couple) for couple in rec_len_bytes_couples[::-1]])
            # Convert from hex to decimal value
            text_length = int(rec_len_bytes, 16)
            text_start_index = start_index + 8
            hex_text = hex_data[text_start_index:text_start_index+text_length*2]
            byte_string = bytes.fromhex(hex_text)            
            result = byte_string.decode(ENCODING, errors="ignore")

            parsed_ppt_dict["content"][str(text_counter)] = result
            text_counter += 1
        
        return parsed_ppt_dict

    except Exception as ex:
        print(f"Something went wrong: {type(ex).__name__}")
        return {}



if __name__ == "__main__":
    args = process_args()
    parsed_ppt_dict = process(args.ppt)
    os.makedirs(args.output_dir, exist_ok=True)
    with open(os.path.join(args.output_dir, 'parsed_ppt.json'), 'w', encoding='utf-8') as f_out:
        json.dump(parsed_ppt_dict, f_out, indent=4, ensure_ascii=False)
