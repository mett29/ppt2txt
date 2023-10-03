import os
import json
from ppt2txt import ppt2txt


def main() -> None:
    args = ppt2txt.process_args()
    os.makedirs(args.output_dir, exist_ok=True)
    parsed_ppt_dict = ppt2txt.process(args.ppt)
    with open(os.path.join(args.output_dir, 'parsed_ppt.json'), 'w', encoding='utf-8') as f_out:
        json.dump(parsed_ppt_dict, f_out, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    main()