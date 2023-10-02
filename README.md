# ppt2txt

A pure python based utility to extract text from PPT files.

The code is based on the official documentation for MS-PPT files available at [https://msopenspecs.azureedge.net/files/MS-PPT/%5bMS-PPT%5d.pdf](https://msopenspecs.azureedge.net/files/MS-PPT/%5bMS-PPT%5d.pdf).

## How to run?

```python
import ppt2txt

# extract content
parsed_ppt_dict = ppt2txt.parse_ppt("file.ppt") 
```

## Output

`parsed_ppt_dict` is a dictionary with the following structure:

```json
{
    "filename": "file_example_PPT_250kB_v2.ppt",
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