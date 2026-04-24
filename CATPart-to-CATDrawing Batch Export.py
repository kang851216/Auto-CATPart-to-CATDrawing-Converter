import os
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas
from win32com.client import Dispatch

DEFAULT_SHEET_NAME = "Sheet"

def get_scale(length_value):
    scale_factor = int(length_value / 310)
    if scale_factor <= 1:
        return 1
    if scale_factor <= 2:
        return 1 / 2
    if scale_factor <= 3:
        return 1 / 3
    if scale_factor <= 5:
        return 1 / 5
    if scale_factor <= 8:
        return 1 / 8
    if scale_factor <= 10:
        return 1 / 10
    if scale_factor <= 12:
        return 1 / 12
    if scale_factor <= 15:
        return 1 / 15
    if scale_factor <= 18:
        return 1 / 18
    if scale_factor <= 20:
        return 1 / 20
    if scale_factor <= 25:
        return 1 / 25
    if scale_factor <= 30:
        return 1 / 30
    return 1 / 40


def build_path(folder, name, extension):
    return os.path.join(folder, f"{name}{extension}")


def safe_close(document):
    if document is None:
        return
    try:
        document.Close()
    except Exception:
        pass


def requires_unfolded_view(process1_value, process2_value):
    keywords = ("rolling", "bending", "翻滚", "折弯")
    for value in (process1_value, process2_value):
        text = "" if pandas.isna(value) else str(value).strip()
        text_lower = text.lower()
        if any(keyword in text_lower or keyword in text for keyword in keywords):
            return True
    return False


def requires_section_view(material_type_value):
    if pandas.isna(material_type_value):
        return False
    material_text = str(material_type_value).strip().lower()
    return material_text in ("solid bar", "structural steel")


def requires_side_view(material_type_value, process2_value):
    if pandas.isna(material_type_value):
        return False

    material_text = str(material_type_value).strip().lower()
    process2_text = "" if pandas.isna(process2_value) else str(process2_value).strip().lower()

    return material_text == "sheet metal" and process2_text in ("rolling", "bending")


def add_section_view(drawing_views, drawing_view_front, product, scale, status_callback, row_index):
    parent_behavior = drawing_view_front.GenerativeBehavior

    # Major axis: horizontal cutting line through front view center.
    # Minor axis: vertical cutting line through front view center.
    section_specs = [
        {
            "name": "major-axis",
            "x": 315,
            "y": 95,
            "profile": [10.0, 0.0, 110.0, 0.0],
        },
        {
            "name": "minor-axis",
            "x": 380,
            "y": 95,
            "profile": [60.0, -40.0, 60.0, 40.0],
        },
    ]

    for section_spec in section_specs:
        drawing_view_section = drawing_views.Add("AutomaticNaming")
        drawing_view_section.X = section_spec["x"]
        drawing_view_section.Y = section_spec["y"]
        drawing_view_section.Scale = scale

        behavior_section = drawing_view_section.GenerativeBehavior
        behavior_section.Document = product

        try:
            behavior_section.DefineSectionView(
                section_spec["profile"],
                "SectionView",
                "Offset",
                1,
                parent_behavior,
            )
        except Exception:
            try:
                behavior_section.DefineSectionView(
                    tuple(section_spec["profile"]),
                    "SectionView",
                    "Offset",
                    1,
                    parent_behavior,
                )
            except Exception as section_error:
                status_callback(
                    f"Row {row_index}: {section_spec['name']} section skipped -> {section_error}"
                )
                continue

        behavior_section.Update()
        status_callback(f"Row {row_index}: {section_spec['name']} section view added")


def run_drawing_generation(config, status_callback):
    info = pandas.read_excel(config["excel_file"], sheet_name=DEFAULT_SHEET_NAME)

    required_columns = ["Part Name", "Quantity", "Material Type", "Process1", "Process2"]
    missing_columns = [col for col in required_columns if col not in info.columns]
    if missing_columns:
        status_callback(
            "BOM format error: missing required column(s): " + ", ".join(missing_columns)
        )
        return

    if len(info) == 0:
        status_callback("No data rows found. BOM must contain data starting from A2.")
        return

    # In pandas, index 0 corresponds to Excel A2 when row 1 is used as header.
    part_name_mask = (
        info["Part Name"].notna()
        & info["Part Name"].astype(str).str.strip().ne("")
    )

    if not part_name_mask.any():
        status_callback("No valid rows detected in column 'Part Name' from A2 onward.")
        return

    last_data_idx = part_name_mask[part_name_mask].index.max()
    process_rows = [row_idx for row_idx in range(0, last_data_idx + 1) if part_name_mask.iloc[row_idx]]

    if not process_rows:
        status_callback("No valid rows detected in BOM data range.")
        return

    status_callback(
        f"Auto-detected row range from A2 to A{last_data_idx + 2} (Excel row numbering)."
    )

    CATIA = Dispatch("CATIA.Application")
    CATIA.Visible = True
    success_count = 0
    error_count = 0

    for i in process_rows:
        part_doc = None
        drawing_document = None
        partname = str(info.at[i, "Part Name"]).strip()
        if not partname or partname == "nan":
            status_callback(f"Row {i}: skipped (empty part name)")
            continue

        try:
            quantity = int(info.at[i, "Quantity"])
        except Exception:
            quantity = 0

        process1 = info.at[i, "Process1"]
        process2 = info.at[i, "Process2"]
        material_type = info.at[i, "Material Type"]
        model = f"{partname}.CATPart"
        part_file = build_path(config["part_path"], partname, ".CATPart")

        if not os.path.exists(part_file):
            status_callback(f"Row {i}: file not found -> {part_file}")
            continue

        try:
            status_callback(f"Row {i}: processing {partname}")

            part_doc = CATIA.Documents.Open(part_file)
            documents = CATIA.Documents
            part_document = documents.Item(model)
            product = part_document.GetItem(partname)

            part_parameters = part_document.Part.Parameters
            try:
                if part_parameters.Item("L").IsTrueParameter:
                    length_value = part_parameters.Item("L").Value
                else:
                    length_value = 1
            except Exception:
                length_value = 1

            scale = get_scale(length_value)

            drawing_document = documents.Add("Drawing")
            drawing_document.Standard = 1
            drawing_sheets = drawing_document.Sheets
            drawing_sheet = drawing_sheets.Item("Sheet.1")
            drawing_sheet.PaperSize = 5
            drawing_views = drawing_sheet.Views

            req1 = "Technical Requirement"
            req2 = "1. Unless otherwise specified, the tolerance is ±0.1mm"
            text1 = drawing_views.ActiveView.Texts.Add(req1, 25, 35)
            text2 = drawing_views.ActiveView.Texts.Add(req2, 25, 25)
            text1.SetFontName(0, 0, "Arial Unicode MS (TrueType)")
            text2.SetFontName(0, 0, "Arial Unicode MS (TrueType)")

            drawing_view_iso = drawing_views.Add("AutomaticNaming")
            drawing_view_iso.X = 315
            drawing_view_iso.Y = 210
            drawing_view_iso.Scale = scale
            behavior_iso = drawing_view_iso.GenerativeBehavior
            behavior_iso.Document = product
            behavior_iso.DefineIsometricView(-0.707, 0.707, 0.707, 0, 0, 0)
            behavior_iso.ColorInheritanceMode = 1
            behavior_iso.RepresentationMode = 0
            behavior_iso.Update()

            drawing_view_front = drawing_views.Add("AutomaticNaming")
            drawing_view_front.X = 210
            drawing_view_front.Y = 148.5
            drawing_view_front.Scale = scale
            behavior_front = drawing_view_front.GenerativeBehavior
            behavior_front.Document = product
            behavior_front.DefineFrontView(1, 0, 0, 0, 1, 0)
            behavior_front.HiddenLineMode = 1
            behavior_front.Update()

            if requires_side_view(material_type, process2):
                drawing_view_side = drawing_views.Add("AutomaticNaming")
                drawing_view_side.X = 105
                drawing_view_side.Y = 148.5
                drawing_view_side.Scale = scale
                behavior_side = drawing_view_side.GenerativeBehavior
                behavior_side.Document = product
                # Side view projection on YZ plane.
                behavior_side.DefineFrontView(0, 1, 0, 0, 0, 1)
                behavior_side.HiddenLineMode = 1
                behavior_side.Update()
                status_callback(f"Row {i}: side view added")

            if requires_unfolded_view(process1, process2):
                drawing_view_unfold = drawing_views.Add("AutomaticNaming")
                drawing_view_unfold.X = 315
                drawing_view_unfold.Y = 148.5
                drawing_view_unfold.Scale = scale
                behavior_unfold = drawing_view_unfold.GenerativeBehavior
                behavior_unfold.Document = product
                behavior_unfold.DefineUnfoldedView(0.0, 0.0, 1.0, 1.0, 0.0, 0.0)
                behavior_unfold.Update()

            if requires_section_view(material_type):
                add_section_view(
                    drawing_views,
                    drawing_view_front,
                    product,
                    scale,
                    status_callback,
                    i,
                )

            time.sleep(2)
            if config["catscript_path"]:
                os.startfile(config["catscript_path"])
                time.sleep(5)

            quantity_text = drawing_views.ActiveView.Texts.Add(str(quantity), 314, 41)
            quantity_text.SetFontName(0, 0, "Arial Unicode MS (TrueType)")

            drawing_sheet.GenerateDimensions
            drawing_view_front.Activate()

            drawing_filename = os.path.join(config["part_path"], partname)
            drawing_document.SaveAs(drawing_filename)
            success_count += 1
            status_callback(f"Row {i}: completed {partname}")
        except Exception as exc:
            error_count += 1
            status_callback(f"Row {i}: error in {partname} -> {exc}")
            status_callback(traceback.format_exc())
        finally:
            safe_close(drawing_document)
            safe_close(part_doc)

    status_callback(f"Completed: {success_count} succeeded, {error_count} failed")


def browse_file(var, title, filetypes):
    selected = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if selected:
        var.set(selected)


def browse_folder(var, title):
    selected = filedialog.askdirectory(title=title)
    if selected:
        var.set(selected)


def main():
    root = tk.Tk()
    root.title("Automatic CATPart-to-CATDrawing Batch Converter")
    root.geometry("860x640")

    excel_var = tk.StringVar(value="")
    part_path_var = tk.StringVar(value="")
    catscript_var = tk.StringVar(value="")

    frame = tk.Frame(root, padx=12, pady=12)
    frame.pack(fill="both", expand=True)

    explanation_text = (
        "1. This is the script for batch process to generate CATDrawing draft file from CATPart file.\n"
        "2. Total three types of files are essential: \n"
        "  (1) BOM.xlsx (Please see the template and customize it based on the template format)\n"
        "  (2) CATDrawing Script file (This is for automatic generation of drawing frame)\n"
        "  (3) Target CATPart file(s)\n"
        "3. How to use: \n"
        "  (1) Select 'BOM' file in xlsx format. \n"
        "  (2) Sheet name is fixed to 'Sheet' in code.\n"
        "  (3) Select 'CATScript' file for automatic generation of drawing frame. \n"
        "  (4) Select the path of folder where CATPart(s) are saved. \n"
        "  (5) Run 'CATIA'. \n"
        "  (6) Click 'Run' button on UI.\n"
        "  (7) Wait until the process is completed. Warning or error messages may appear in CATIA, but the process will continue.\n"
        "  (8) The generated CATDrawing file will be saved in the same folder of CATPart file with same name. \n"
    )
    explanation_label = tk.Label(
        frame,
        text=explanation_text,
        justify="left",
        anchor="w",
        wraplength=810,
        bg="#f3f3f3",
        relief="groove",
        padx=10,
        pady=10,
    )
    explanation_label.grid(row=0, column=0, columnspan=3, sticky="we", pady=(0, 12))

    def add_row(row, label, var, browse_cmd=None):
        tk.Label(frame, text=label, anchor="w").grid(row=row, column=0, sticky="w", pady=4)
        tk.Entry(frame, textvariable=var, width=70).grid(row=row, column=1, sticky="we", pady=4)
        if browse_cmd:
            tk.Button(frame, text="Browse", command=browse_cmd).grid(row=row, column=2, padx=6)

    add_row(
        1,
        "BOM Excel file",
        excel_var,
        lambda: browse_file(excel_var, "Select Excel file", [("Excel files", "*.xlsx;*.xls")]),
    )
    add_row(
        2,
        "CATScript",
        catscript_var,
        lambda: browse_file(catscript_var, "Select CATScript", [("CATScript", "*.CATScript"), ("All files", "*.*")]),
    )
    add_row(3, "Part folder", part_path_var, lambda: browse_folder(part_path_var, "Select part folder"))

    frame.grid_columnconfigure(1, weight=1)

    status_box = tk.Text(frame, height=10, width=95, state="disabled")
    status_box.grid(row=5, column=0, columnspan=3, pady=10, sticky="nsew")
    frame.grid_rowconfigure(5, weight=1)

    def append_status(message):
        status_box.config(state="normal")
        status_box.insert("end", f"{message}\n")
        status_box.see("end")
        status_box.config(state="disabled")

    def run_clicked():
        config = {
            "excel_file": excel_var.get().strip(),
            "part_path": part_path_var.get().strip(),
            "catscript_path": catscript_var.get().strip(),
        }

        if not os.path.exists(config["excel_file"]):
            messagebox.showerror("Input error", "Excel file not found.")
            return
        if not os.path.isdir(config["part_path"]):
            messagebox.showerror("Input error", "Part folder not found.")
            return

        run_button.config(state="disabled")
        append_status("Starting drawing generation...")

        def worker():
            try:
                run_drawing_generation(config, lambda msg: root.after(0, append_status, msg))
                root.after(0, messagebox.showinfo, "Finished", "Drawing generation completed.")
            except Exception as exc:
                root.after(0, messagebox.showerror, "Error", str(exc))
            finally:
                root.after(0, run_button.config, {"state": "normal"})

        threading.Thread(target=worker, daemon=True).start()

    run_button = tk.Button(frame, text="Run", width=16, command=run_clicked)
    run_button.grid(row=4, column=0, pady=6, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    main()

