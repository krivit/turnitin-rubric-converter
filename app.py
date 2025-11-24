from shiny import App, ui, render, reactive, req
import tempfile
import os
import json

from rubric_converter import excel_to_rbc, rbc_to_excel, excel_to_ims, is_ims_format

def get_output_filename(input_name, output_format=None):
    base, ext = os.path.splitext(input_name)
    if ext.lower() in [".rbc", ".json"]:
        return base + ".xlsx"
    elif ext.lower() == ".xlsx":
        if output_format == "ims":
            return base + ".json"
        else:
            return base + ".rbc"
    else:
        raise ValueError("Unsupported file type: " + ext)

def create_example_json(filepath):
    rubric = [{
        "total_points": None,
        "criterion": [2000000, 2000001],
        "id": 1,
        "scoring_method": 4,
        "name": "Example Rubric",
        "distribute_criterion_percentage": 0,
        "rubric_group": None,
        "is_starred": 0,
        "deleted": 0,
        "criterion_scales_all": [
            3000000, 3000001, 3000002, 3000003,
            3000004, 3000005, 3000006, 3000007
        ],
        "scale_values": [1000000, 1000001, 1000002, 1000003],
        "papers_scored": 0,
        "owner": 0,
        "cv_loaded": "1",
        "description": None
    }]
    scales = [
        {"id": 1000000, "num": 1, "position": 1, "value": 0, "name": "Excellent", "rubric": 1},
        {"id": 1000001, "num": 2, "position": 2, "value": 0, "name": "Good", "rubric": 1},
        {"id": 1000002, "num": 3, "position": 3, "value": 0, "name": "Needs Improvement", "rubric": 1},
        {"id": 1000003, "num": 4, "position": 4, "value": 0, "name": "Absent", "rubric": 1},
    ]
    criteria = [
        {
            "value": 0,
            "id": 2000000,
            "rubric": 1,
            "name": "Analysis",
            "description": "Quality and depth of analysis.",
            "criterion_scales": [3000000, 3000001, 3000002, 3000003],
            "position": 1,
            "previous_version": None,
            "num": 1
        },
        {
            "value": 0,
            "id": 2000001,
            "rubric": 1,
            "name": "Writing",
            "description": "Clarity and structure.",
            "criterion_scales": [3000004, 3000005, 3000006, 3000007],
            "position": 2,
            "previous_version": None,
            "num": 2
        }
    ]
    criterion_scales = [
        {"criterion": 2000000, "scale_value": 1000000, "description": "Insightful and thorough", "value": 5, "id": 3000000},
        {"criterion": 2000000, "scale_value": 1000001, "description": "Adequate analysis", "value": 4, "id": 3000001},
        {"criterion": 2000000, "scale_value": 1000002, "description": "Superficial", "value": 2, "id": 3000002},
        {"criterion": 2000000, "scale_value": 1000003, "description": None, "value": 0, "id": 3000003},
        {"criterion": 2000001, "scale_value": 1000000, "description": "Clear, well-organised", "value": 3, "id": 3000004},
        {"criterion": 2000001, "scale_value": 1000001, "description": "Generally clear", "value": 2, "id": 3000005},
        {"criterion": 2000001, "scale_value": 1000002, "description": "Difficult to follow", "value": 1, "id": 3000006},
        {"criterion": 2000001, "scale_value": 1000003, "description": None, "value": 0, "id": 3000007},
    ]
    example = {
        "Rubric": rubric,
        "RubricCriterion": criteria,
        "RubricScale": scales,
        "RubricCriterionScale": criterion_scales
    }
    with open(filepath, "w") as f:
        json.dump(example, f, indent=2)

def convert_json_to_excel(json_path, excel_path):
    rbc_to_excel(json_path, excel_path)

app_ui = ui.page_fluid(
    ui.h2("Turnitin/IMS Rubric Converter"),
    ui.tags.style("""
    .compact-download-btn {
        font-size: 0.85em;
        padding: 2px 10px;
        margin-bottom: 2px;
        margin-right: 8px;
        vertical-align: middle;
        background-color: #f8f8f8;
        border: 1px solid #bbb;
        border-radius: 4px;
        cursor: pointer;
        display: inline-block;
    }
    """),
    ui.markdown("#### Creating a new rubric"),
    ui.tags.ol(
        ui.tags.li(
            ui.download_button("examplefile", "Download the Example Excel File.", class_="compact-download-btn")
        ),
        ui.tags.li("Proceed from Step 3 below."),
    ),
    ui.markdown(
        """
#### Editing an existing rubric
1. Download the rubric as a `.rbc` file (Turnitin format) or `.json` file (IMS format).
    - For Turnitin: The Download option is available in the sandwich menu — the three lines menu — for the rubric. 
2. Upload the `.rbc` or `.json` file below for conversion, and download the resulting Excel (`.xlsx`) file.
3. Edit the Excel file.
    - To specify a criterion title and description: put the title on the first line and the description on the second line of the cell in the "Criterion (name and description)" column.
    - Editing the column names in the first row will edit the scale names. (Note that IMS ignores them.)
    - For cell descriptions and values: use "desc [value]", "[value]", or just "desc".
    - Note that Turnitin requires all criteria to have the same number of cells, though their point values may differ.
4. Upload the `.xlsx` file below, select the output format (Turnitin or IMS), and download the resulting file.
    - You can optionally edit the rubric name before downloading.
5. Upload the file back to your platform.
    - For Turnitin: The Upload option is available in the sandwich menu for the rubric.
"""
    ),
    ui.tags.div(
        ui.tags.strong(
            ui.input_file("file", "Upload .rbc, .json, or .xlsx file for conversion here.", accept=[".rbc", ".xlsx", ".json"]),
            style="font-size: 1.13em;"
        ),
        style="margin-bottom: 8px; margin-top: 18px;"
    ),
    ui.output_text_verbatim("status"),
    ui.output_ui("format_selector_ui"),
    ui.output_ui("rubric_name_ui"),
    ui.output_ui("download_ui"),
)

def server(input, output, session):
    converted_file = reactive.Value(None)
    converted_name = reactive.Value(None)
    status_message = reactive.Value("")
    uploaded_excel = reactive.Value(None)
    uploaded_excel_name = reactive.Value(None)
    rubric_name_value = reactive.Value("")
    output_format = reactive.Value("turnitin")

    @reactive.Effect
    @reactive.event(input.file)
    def _():
        fileinfo = input.file()
        if not fileinfo:
            return
        in_path = fileinfo[0]["datapath"]
        in_name = fileinfo[0]["name"]
        ext = os.path.splitext(in_name)[1].lower()
        if ext not in [".rbc", ".xlsx", ".json"]:
            status_message.set("Unsupported file type.")
            converted_file.set(None)
            converted_name.set(None)
            uploaded_excel.set(None)
            uploaded_excel_name.set(None)
            rubric_name_value.set("")
            output_format.set("turnitin")
            return
        outname = get_output_filename(in_name)
        outpath = os.path.join(tempfile.gettempdir(), outname)
        if os.path.exists(outpath):
            os.remove(outpath)
        try:
            if ext in [".rbc", ".json"]:
                rbc_to_excel(in_path, outpath)
                status_message.set(f"Conversion successful. Output: {outname}")
                converted_file.set(outpath)
                converted_name.set(outname)
                uploaded_excel.set(None)
                uploaded_excel_name.set(None)
                rubric_name_value.set("")
                output_format.set("turnitin")
            elif ext == ".xlsx":
                uploaded_excel.set(in_path)
                uploaded_excel_name.set(in_name)
                base = os.path.splitext(in_name)[0].replace("_", " ")
                rubric_name_value.set(base)
                status_message.set("Excel file uploaded. Select output format, optionally edit rubric name, then convert.")
                converted_file.set(None)
                converted_name.set(None)
                output_format.set("turnitin")
        except Exception as e:
            status_message.set(f"Conversion failed: {e}")
            converted_file.set(None)
            converted_name.set(None)
            uploaded_excel.set(None)
            uploaded_excel_name.set(None)
            rubric_name_value.set("")
            output_format.set("turnitin")

    @output
    @render.text
    def status():
        return status_message.get()

    @output
    @render.ui
    def format_selector_ui():
        if uploaded_excel.get():
            return ui.input_radio_buttons(
                "output_format", 
                "Output format:", 
                {"turnitin": "Turnitin (.rbc)", "ims": "IMS (.json)"},
                selected="turnitin"
            )
        return ""

    @output
    @render.ui
    def rubric_name_ui():
        if uploaded_excel.get():
            return ui.input_text("rubricname", "Rubric name (edit before converting):", value=rubric_name_value.get())
        return ""

    @output
    @render.ui
    def download_ui():
        if converted_file.get():
            return ui.download_button("downloadfile", "Download converted file")
        return ""

    @output(id="examplefile")
    @render.download(filename="example_rubric.xlsx")
    def examplefile():
        tmpdir = tempfile.gettempdir()
        json_path = os.path.join(tmpdir, "example_rubric.json")
        excel_path = os.path.join(tmpdir, "example_rubric.xlsx")
        create_example_json(json_path)
        convert_json_to_excel(json_path, excel_path)
        with open(excel_path, "rb") as f:
            yield f.read()

    @reactive.Effect
    @reactive.event(input.rubricname, input.output_format, uploaded_excel)
    def _():
        if uploaded_excel.get() and rubric_name_value.get():
            in_path = uploaded_excel.get()
            in_name = uploaded_excel_name.get()
            rubric_name = rubric_name_value.get()
            fmt = input.output_format() if input.output_format() else "turnitin"
            output_format.set(fmt)
            outname = get_output_filename(in_name, fmt)
            outpath = os.path.join(tempfile.gettempdir(), outname)
            if os.path.exists(outpath):
                os.remove(outpath)
            try:
                if fmt == "ims":
                    excel_to_ims(in_path, outpath, rubric_name_override=rubric_name)
                else:
                    excel_to_rbc(in_path, outpath, rubric_name_override=rubric_name)
                status_message.set(f"Conversion successful. Output: {outname}")
                converted_file.set(outpath)
                converted_name.set(outname)
            except Exception as e:
                status_message.set(f"Conversion failed: {e}")
                converted_file.set(None)
                converted_name.set(None)

    @reactive.Effect
    @reactive.event(input.rubricname)
    def _():
        if input.rubricname() is not None:
            rubric_name_value.set(input.rubricname())

    @output(id="downloadfile")
    @render.download(filename=lambda: converted_name.get() or "converted")
    def downloadfile():
        ofile = converted_file.get()
        if ofile and os.path.exists(ofile):
            with open(ofile, "rb") as f:
                yield f.read()

app = App(app_ui, server)
