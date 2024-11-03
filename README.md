# BOLD XLSX

## About

This is a program written in Python3 that adds bold styling to .xlsx files.

More specifically, it applies bold styling to \*\*bold\*\* markdown syntax found in cells.

This includes cells that reference other cells (even in other worksheets), since in such cases Excel does not transfer the styling along with the cell value.

However, in such cases the reference is lost and the cell becomes a normal text cell. For this reason, the output is a separate .xlsx file, and the original file should remain intact.

## Disclaimer

Since this program was created to account for one very specific use case there may be situations unaccounted for.

That is to say, you should probably make a backup of your .xlsx file before running this program on it.

## Dependencies

lxml==5.3.0 \[[WebSite](https://lxml.de/)\]


## Usage

1. Clone this repository or simply download the bold_xlsx.py script
2. Install the lxml dependency
3. Run the bold_xlsx.py script like so:

```bash
>python3 bold_xlsx.py [path_to_xlsx_file]
```

4. The output file will be saved to the same directory as the input file, with the name:

```
BOLD_{original_name}.xlsx
```


