# MHW Decorations Export - Comparison Tool
Tailored for [MHW Decorations Export Files](https://github.com/yNEX/mhw_decos_csv_advanced), this tool streamlines the comparison of differences between exports. Quickly discern variations, and export the results to Excel, text, or a command-line table. It highlights new decorations and shows how many are added for existing ones, together with the quantity.

# Usage
You just need to download the `decoCompare.py` and `requirements.txt`file.<br>
After installing Python run `pip install -r {path to requirements.txt}`.<br><br>
Now you can run the script in your terminal by running `python decoCompare.py`
```
Usage: decoCompare.py [options] <old_export> <new_export>

Compares two export files (JSON or TXT) used for tracking changes in decorations in Monster Hunter World.
This tool allows tracking changes in decorations and formats the output for Excel or as text.

Options:
  -e, --excel    Path and filename for Excel output. Defaults to 'DecoChanges.xlsx' in the current directory
  -t, --text     Path and filename for Text output. Defaults to 'DecoChanges.txt' in the current directory
  -b, --both     Create both Excel and Text versions. Defaults to 'DecoChanges.xlsx' and 'DecoChanges.txt'
  -h, --help     Show this help message

By default, outputs are displayed in the terminal.
```
# Requirements
-  [Python 3.9+](https://www.python.org/downloads/)
-  pip (comes with Python)
-  Python packages listed in `requirements.txt`
# Todo
- [ ] Finish localization
- [ ] Create an executable and distribute
- [ ] rewrite in Java
- [x] ~Change formula for sum of newly added decos~
- [x] ~File opening when pressing just {Enter} instead of providing an option~
