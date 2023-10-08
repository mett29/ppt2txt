# ppt2txt

A pure python based utility to extract text from PPT files.

The code is based on the official documentation for MS-PPT files available at [https://msopenspecs.azureedge.net/files/MS-PPT/%5bMS-PPT%5d.pdf](https://msopenspecs.azureedge.net/files/MS-PPT/%5bMS-PPT%5d.pdf).

## How to install?

```bash
pip install ppt2txt
```

## How to run?

- From command line:
```bash
ppt2txt file.ppt -o output_dir
```

- From python:
```python
import ppt2txt

# extract content
parsed_ppt_dict = ppt2txt.process("file.ppt") 
```

## Output

`parsed_ppt_dict` is a dictionary with the following structure:

```json
{
    "filename": "file.ppt",
    "slides": 4,
    "content": {
        "0": "Text from the first record",
        "1": "Text from the second record"
    }
}
```

where:
- `filename` is the name of the input file
- `slides` is the number of slides
- `content` is a dictionary containing an element for each record of type text found in the document