from pathlib import Path

import wx
from openpyxl import load_workbook


def export_budget(workbook_path, output_folder):
    workbook = load_workbook(workbook_path, data_only=True)

    if "budget setup" not in workbook.sheetnames:
        raise ValueError(
            'The selected workbook does not contain a "budget setup" worksheet.'
        )

    worksheet = workbook["budget setup"]
    output_rows = []

    for row in range(18, 29):
        category = worksheet[f"A{row}"].value

        if not category:
            continue

        category = str(category)

        if "-" not in category:
            raise ValueError(
                f'Cell A{row} in "budget setup" must contain a category '
                'in the format "number - description".'
            )

        catnum, catdes = category.split("-", 1)

        hours1 = worksheet[f"B{row}"].value or 0
        cost1 = worksheet[f"D{row}"].value or 0
        cost5 = worksheet[f"E{row}"].value or 0
        cost6 = worksheet[f"F{row}"].value or 0
        cost7 = worksheet[f"G{row}"].value or 0
        cost8 = worksheet[f"H{row}"].value or 0
        cost9 = worksheet[f"I{row}"].value or 0
        cost10 = 0
        cost11 = worksheet[f"J{row}"].value or 0
        cost12 = worksheet[f"K{row}"].value or 0
        cost14 = worksheet[f"L{row}"].value or 0

        output_rows.append([
            catnum.strip(),
            catdes.strip(),
            hours1,
            cost1,
            cost5,
            cost6,
            cost7,
            cost8,
            cost9,
            cost10,
            cost11,
            cost12,
            cost14,
        ])

    if not output_rows:
        raise ValueError("No budget categories were found in rows 18 through 28.")

    output_path = Path(output_folder) / "ceBudget.txt"

    with output_path.open("w", encoding="utf-8", newline="") as file:
        for row in output_rows:
            file.write("\t".join(str(value) for value in row) + "\n")

    return output_path, len(output_rows)


class BudgetExportApp(wx.Frame):
    def __init__(self):
        super().__init__(
            None,
            title="ceBudget Export",
            size=(650, 230),
        )

        panel = wx.Panel(self)
        layout = wx.BoxSizer(wx.VERTICAL)

        self.workbook_picker = wx.FilePickerCtrl(
            panel,
            message="Select the Excel workbook",
            wildcard="Excel workbooks (*.xlsx)|*.xlsx",
            style=wx.FLP_OPEN | wx.FLP_FILE_MUST_EXIST,
        )

        self.output_picker = wx.DirPickerCtrl(
            panel,
            message="Select the output folder",
            style=wx.DIRP_USE_TEXTCTRL,
        )

        layout.Add(
            wx.StaticText(panel, label="Excel workbook:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.workbook_picker,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        layout.Add(
            wx.StaticText(panel, label="Save ceBudget.txt to:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.output_picker,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        export_button = wx.Button(panel, label="Export Budget")
        export_button.Bind(wx.EVT_BUTTON, self.on_export)
        layout.Add(export_button, 0, wx.ALL | wx.ALIGN_RIGHT, 10)

        panel.SetSizer(layout)
        self.Centre()

    def on_export(self, event):
        workbook_path = self.workbook_picker.GetPath()
        output_folder = self.output_picker.GetPath()

        if not workbook_path or not output_folder:
            wx.MessageBox(
                "Select an Excel workbook and an output folder.",
                "Missing selection",
                wx.OK | wx.ICON_WARNING,
            )
            return

        try:
            output_path, row_count = export_budget(
                workbook_path,
                output_folder,
            )

            wx.MessageBox(
                f"Created ceBudget.txt with {row_count} row(s):\n{output_path}",
                "Export complete",
                wx.OK | wx.ICON_INFORMATION,
            )
        except Exception as error:
            wx.MessageBox(
                f"Budget export failed:\n{error}",
                "Error",
                wx.OK | wx.ICON_ERROR,
            )


if __name__ == "__main__":
    app = wx.App(False)
    frame = BudgetExportApp()
    frame.Show()
    app.MainLoop()
