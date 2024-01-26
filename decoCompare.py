import sys
import json
import pandas as pd
import argparse
import os
from prettytable import PrettyTable
import gettext
import locale

# Set up gettext for internationalization
current_locale, _ = locale.getlocale()
gettext.bindtextdomain('base', 'locales')
gettext.textdomain('base')
_ = gettext.gettext
locale.setlocale(locale.LC_ALL, current_locale)

# Name of the current script
script_name = os.path.basename(__file__)

# Default variables
excel_created = False
text_created = False

# Accepted file extensions
accepted_extensions = ['.json', '.txt']

def custom_help():
    help_text = (
        _("\nUsage: {0} [options] <old_export> <new_export>\n").format(script_name)
        + _("\nCompares two export files (JSON or TXT) used for tracking changes in decorations in Monster Hunter World. ")
        + _("\nThis tool allows tracking changes in decorations and formats the output for Excel or as text. ")
        + _("\nOptions:\n")
        + _("  -e, --excel    Path and filename for Excel output. Defaults to 'DecoChanges.xlsx' in the current directory\n")
        + _("  -t, --text     Path and filename for Text output. Defaults to 'DecoChanges.txt' in the current directory\n")
        + _("  -b, --both     Create both Excel and Text versions. Defaults to 'DecoChanges.xlsx' and 'DecoChanges.txt'\n")
        + _("  -h, --help     Show this help message\n")
        + _("\nBy default, outputs are displayed in the terminal.")
    )
    print(help_text)
    exit()

class CustomArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        custom_help()

def is_file_extension_accepted(file_path):
    return any(file_path.lower().endswith(extension) for extension in accepted_extensions)

def load_export(file_path):
    if not is_file_extension_accepted(file_path):
        raise ValueError(_("\nFile extension not supported. Supported extensions: ") + str(accepted_extensions))

    with open(file_path, 'r') as file:
        content = file.read()
        # Remove the warning line, if present
        if content.startswith('WARNING:'):
            content = content.split('\n', 1)[1]
        return json.loads(content)

def compare_json(file1, file2, excel_format=False, text_format=False, both_format=False):
    jsonObject1 = load_export(file1)
    jsonObject2 = load_export(file2)

    # Check if the lists of key-value pairs in the JSON files are identical
    if sorted(jsonObject1.items()) == sorted(jsonObject2.items()):
        print(_("The JSON files contain identical data. The files will not be compared."))
        sys.exit()

    changes = []
    new_decos = []
    changes_text = []
    new_decos_text = []
    
    # Number of decorations changed and new decorations added
    new_keys = set(jsonObject2) - set(jsonObject1)
    for key in sorted(jsonObject1.keys() | jsonObject2.keys()):
        if key in new_keys or (jsonObject1.get(key, 0) == 0 and jsonObject2[key] > 0):
            # Treat as new decoration
            amount = jsonObject2[key]
            new_decos.append({"Decoration": key, "Amount": amount})
            if not excel_format:
                new_decos_text.append(_("{}, amount: {}").format(key, amount))
        elif key in jsonObject2 and jsonObject1[key] != jsonObject2[key]:
            # Treat as changed decoration
            difference = jsonObject2[key] - jsonObject1[key]
            changes.append({"Decoration": key, "Added": difference, "Total": jsonObject2[key]})
            if not excel_format:
                changes_text.append(_("{}, added: {} | {}").format(key, difference, jsonObject2[key]))


    # Calculation of the total number of added decorations
    total_changed = sum(item['Added'] for item in changes)
    total_new = sum(1 for item in new_decos)


    return changes, new_decos, changes_text, new_decos_text, total_changed, total_new

def print_table(data, header, has_pipe=False):
    table = PrettyTable()
    table.field_names = header
    for row in data:
        if has_pipe:
            # Split the line at the pipe and remove "added:"
            row_data = []
            for part in row.split("|"):
                for elem in part.split(", "):
                    if "added:" in elem:
                        row_data.append(elem.split(": ")[1].strip())
                    else:
                        row_data.append(elem.strip())
        else:
            # Remove the word "Amount:" from the line
            row_data = [elem.split(": ")[1] if ":" in elem else elem for elem in row.split(", ")]

        if len(row_data) == len(header):
            table.add_row(row_data)
    print(table)
    
def ensure_file_extension(filename, default_name, extension):
    if os.path.isdir(filename):
        # If the specified path is a directory, add the default name
        return os.path.join(filename, default_name + extension)
    elif not filename.lower().endswith(extension):
        # If the path does not have an extension, add the extension
        return filename + extension
    return filename

if __name__ == "__main__":
    try:    
        # Check if the script was called with the help argument
        if '-h' in sys.argv or '--help' in sys.argv:
            custom_help()

        parser = CustomArgumentParser(add_help=False)
        parser.add_argument("old_export", help=_("Path to the first JSON file (old decorations)"))
        parser.add_argument("new_export", help=_("Path to the second JSON file (new decorations)"))
        parser.add_argument("-e", "--excel", nargs='?', const="DecoChanges.xlsx", help=_("Path and filename for Excel output (default: DecoChanges.xlsx)"))
        parser.add_argument("-t", "--text", nargs='?', const="DecoChanges.txt", help=_("Path and filename for Text output (default: DecoChanges.txt)"))
        parser.add_argument("-b", "--both", action="store_true", help=_("Create both Excel and Text versions"))

        args = parser.parse_args()      

        changes, new_decos, changes_text, new_decos_text, total_changed, total_new = compare_json(args.old_export, args.new_export)

        # Check and set file extensions
        excel_path = ensure_file_extension(args.excel if args.excel is not None else "DecoChanges", "DecoChanges", ".xlsx")
        text_path = ensure_file_extension(args.text if args.text is not None else "DecoChanges", "DecoChanges", ".txt")

        # Create directories if not present
        for path in [excel_path, text_path]:
            directory = os.path.dirname(path)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)

        # If --both is specified, ensure that both Excel and Text paths are set
        if args.both:
            excel_path = "DecoChanges.xlsx"
            text_path = "DecoChanges.txt"

        # Excel Output
        if args.excel or args.both:
            try:
                excel_created = True
                excel_path = os.path.join(os.getcwd(), excel_path)
                with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                    # Add DataFrames to different sheets
                    if changes:
                        df_changes = pd.DataFrame(changes)
                        # Translate column names
                        df_changes.columns = [_(col) for col in df_changes.columns]
                        # Sort by translated columns and export to excel
                        df_changes = df_changes.sort_values(by=_('Decoration'))
                        df_changes.to_excel(writer, sheet_name=_('Existing Decorations'), index=False)
    
                        # Calculate the sum in Excel for "Existing Decorations"
                        worksheet_existing = writer.sheets[_('Existing Decorations')]
                        last_row = len(df_changes.index)
                        worksheet_existing.write(last_row+2, 0, _('Total number added:'))
                        worksheet_existing.write_formula(last_row+2, 1, f'=SUM(B2:B{last_row+1})')
                        worksheet_existing.write_formula(last_row+2, 2, f'=SUM(C2:C{last_row+1})')
                    else:
                        print(_("\nNo changes to existing decorations compared to previous data. The creation of the 'Existing Decorations' workbook has been skipped."))
                    if new_decos:
                        df_new_decos = pd.DataFrame(new_decos)
                        # Translate column names
                        df_new_decos.columns = [_(col) for col in df_new_decos.columns]
                        # Sort by translated columns and export to excel
                        df_new_decos = df_new_decos.sort_values(by=_('Decoration'))
                        df_new_decos.to_excel(writer, sheet_name=_('New Decorations'), index=False)
    
                        # Calculate the sum in Excel for "New Decorations"
                        worksheet_new = writer.sheets[_('New Decorations')]
                        last_row = len(df_new_decos.index)
                        worksheet_new.write(last_row+2, 0, _('Total number added:'))
                        worksheet_new.write(last_row+3, 0, _('Total number added variations:'))
                        worksheet_new.write_formula(last_row+2, 1, f'=SUM(B2:B{last_row+1})')
                        worksheet_new.write_formula(last_row+3, 1, f'=COUNTA(A2:A{last_row+1})')
                    else:
                        print(_("\nNo new decoration types identified compared to the previous export. The creation of the 'New Decorations' workbook has been skipped."))
    
                    if changes or new_decos:
                        # Set formats for headers and cells
                        header_format = writer.book.add_format({
                            'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bold': True
                        })
                        cell_format = writer.book.add_format({
                            'align': 'left', 'valign': 'vcenter'
                        })
                        # Format table headers and contents for both sheets
                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            df_current = df_changes if sheet_name == _('Existing Decorations') else df_new_decos
                            # Add table without colored rows
                            worksheet.add_table(0, 0, df_current.shape[0], df_current.shape[1] - 1, {
                                'columns': [{'header': column_name, 'header_format': header_format, 'format': cell_format}
                                            for column_name in df_current.columns],
                                'autofilter': True,
                                'style': 'Table Style Light 1',  # Standard Excel table without alternating colors
                            })
                            # Auto-adjust column width and apply cell format
                            for column in df_current.columns:
                                column_length = max(df_current[column].astype(str).map(len).max(), len(column))
                                column_length += 6  # Adding a little extra space for the filter dropdown menu
                                col_idx = df_current.columns.get_loc(column)
                                worksheet.set_column(col_idx, col_idx, column_length, cell_format)
                            # Freeze the first row
                            worksheet.freeze_panes(1, 0)
                        print(_("Comparison data saved in the Excel file '{0}'.").format(excel_path))
                    else:
                        print(_("\nNo changes found between the two specified files. Therefore, no file was created."))
            except PermissionError:
                print(_("Error: Access to '{0}' denied. Ensure that the file is not open in another program and that you have the necessary permissions.").format(excel_path))
                excel_created = False
        # Text Output
        if args.text or args.both:
            try:  
                text_created = True
                text_path = os.path.join(os.getcwd(), text_path)
                with open(text_path, 'w') as file:
                    # Changes to existing decorations
                    if changes_text:
                        file.write(_("-----Changes to Existing Decorations-----\n"))
                        for line in sorted(changes_text):
                            file.write(line + "\n")
                    else:
                        file.write(_("-----No Changes to Existing Decorations-----\n"))
                    # Newly added decorations
                    if new_decos_text:
                        file.write(_("\n-----Newly Added Decorations-----\n"))
                        for line in sorted(new_decos_text):
                            file.write(line + "\n")
                    else:
                        file.write(_("\n-----No Newly Added Decorations-----\n"))
                    # Total added
                    file.write(_("\nTotal added (changed decorations): {0}").format(total_changed))
                    file.write(_("\nTotal added (new decorations): {0}").format(total_new))
                    print(_("Comparison data saved in the text file '{0}'.").format(text_path))
            except PermissionError:
                print(_("Error: Access to '{0}' denied. Ensure that the file is not open in another program and that you have the necessary permissions.").format(text_path))
                text_created = False
        # Terminal output if neither --text nor --excel is specified
        if not (args.text or args.excel or args.both):
            args.excel = "DecoChanges.xlsx"
            args.text = "DecoChanges.txt"
            print(_("Changes to Existing Decorations:"))
            sorted_changes = sorted(changes_text, key=lambda x: x.split(", ")[0])
            print_table(sorted_changes, [_("Decoration"), _("Added"), _("Total")], has_pipe=True)
            print(_("\nNewly Added Decorations:"))
            sorted_new_decos = sorted(new_decos_text, key=lambda x: x.split(", ")[0])
            print_table(sorted_new_decos, [_("Decoration"), _("Amount")])
            print(_("\nTotal number added (changed decorations): {0}").format(total_changed))
            print(_("\nTotal number added (new decorations): {0}").format(total_new))
            exit()

        # User interaction only for created files
        if excel_created or text_created:
            print(_("\nPress Enter to open the created files, or choose one of the options:"))
            if excel_created and text_created:
                print(_("[e] Excel\n[t] Text\n[q] Exit"))
            elif excel_created:
                print(_("[e] Excel\n[q] Exit"))
            elif text_created:
                print(_("[t] Text\n[q] Exit"))
            user_input = input(_("Your choice: ")).lower()
            if user_input == 'e' and excel_created:
                os.startfile(excel_path)
            elif user_input == 't' and text_created:
                os.startfile(text_path)
            elif user_input == '':
                # Open both files if both were created
                if excel_created:
                    os.startfile(excel_path)
                if text_created:
                    os.startfile(text_path)
            elif user_input == 'q':
                print(_("Exiting script."))
                exit()
            else:
                print(_("Invalid input or file not created."))

    except Exception as e:
        print(_("An unexpected error occurred: \n{0}").format(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)