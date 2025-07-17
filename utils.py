import os
import shutil
from openpyxl.utils import get_column_letter

# pylint: disable=protected-access


def clear_folder(path: str) -> None:
    """Clears the contents of a folder. Creates the folder if it does not exist."""
    if not os.path.exists(path):
        os.makedirs(path)
        return

    for filename in os.listdir(path):
        file_path = os.path.join(path, filename)
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path)


def get_cell_position(anchor) -> tuple[int, int, str]:
    """Returns (col_num, row_num, cell_label) from anchor object."""
    col = anchor._from.col + 1
    row = anchor._from.row + 1
    cell = f"{get_column_letter(col)}{row}"
    return col, row, cell
