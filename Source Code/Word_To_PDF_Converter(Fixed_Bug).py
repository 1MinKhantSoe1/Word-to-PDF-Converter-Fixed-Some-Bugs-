from tkinter import *
from tkinter import filedialog
from pathlib import Path
import subprocess
import sys
import json
from tqdm.auto import tqdm
from tkinter.messagebox import showinfo


def main():
    def openLocation():

        global Location_Name

        try:

            Location_Name = filedialog.askopenfile(filetypes=[('word file', '*.docx')])

            locationError.config(text=Location_Name.name, fg="green")

        except AttributeError:

            showinfo("Warning!", "Please choose the file that you want to convert !!!")

    def windows(paths, keep_active):
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        wdFormatPDF = 17

        if paths["batch"]:
            for docx_filepath in tqdm(sorted(Path(paths["input"]).glob("*.docx"))):
                pdf_filepath = Path(paths["output"]) / (str(docx_filepath.stem) + ".pdf")
                doc = word.Documents.Open(str(docx_filepath))
                doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
                doc.Close()
        else:
            pbar = tqdm(total=1)
            docx_filepath = Path(paths["input"]).resolve()
            pdf_filepath = Path(paths["output"]).resolve()
            doc = word.Documents.Open(str(docx_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
            pbar.update(1)

        if not keep_active:
            word.Quit()

    def resolve_paths(input_path, output_path):
        input_path = Path(input_path).resolve()
        output_path = Path(output_path).resolve() if output_path else None
        output = {}
        if input_path.is_dir():
            output["batch"] = True
            output["input"] = str(input_path)
            if output_path:
                assert output_path.is_dir()
            else:
                output_path = str(input_path)
            output["output"] = output_path
        else:
            output["batch"] = False
            assert str(input_path).endswith(".docx")
            output["input"] = str(input_path)
            if output_path and output_path.is_dir():
                output_path = str(output_path / (str(input_path.stem) + ".pdf"))
            elif output_path:
                assert str(output_path).endswith(".pdf")
            else:
                output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
            output["output"] = output_path
        return output

    def macos(paths, keep_active):
        script = (Path(__file__).parent / "convert.jxa").resolve()
        cmd = [
            "/usr/bin/osascript",
            "-l",
            "JavaScript",
            str(script),
            str(paths["input"]),
            str(paths["output"]),
            str(keep_active).lower(),
        ]

        def run(cmd):
            process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
            while True:
                line = process.stderr.readline().rstrip()
                if not line:
                    break
                yield line.decode("utf-8")

        total = len(list(Path(paths["input"]).glob("*.docx"))) if paths["batch"] else 1
        pbar = tqdm(total=total)
        for line in run(cmd):
            try:
                msg = json.loads(line)
            except ValueError:
                continue
            if msg["result"] == "success":
                pbar.update(1)
            elif msg["result"] == "error":
                print(msg)
                sys.exit(1)

    def convert(input_path, output_path=None, keep_active=False):
        paths = resolve_paths(input_path, output_path)
        if sys.platform == "darwin":
            return macos(paths, keep_active)
        elif sys.platform == "win32":
            return windows(paths, keep_active)
        else:
            raise NotImplementedError(
                "docx2pdf is not implemented for linux as it requires Microsoft Word to be installed"
            )

    def c():

        try:

            convert(Location_Name.name)

            showinfo("Done", "Successfully Converted!!")

        except AttributeError:

            showinfo("Warning!", "You need to choose the file first !!!")

    tk = Tk()

    tk.title("Word to PDF Converter")
    tk.geometry("350x400")
    tk.columnconfigure(0, weight=1)

    blank = Label()
    blank.grid()

    blank = Label()
    blank.grid()

    blank = Label()
    blank.grid()

    # Location Label
    locationLabel = Label(tk, text="Please Select The Docx File That You Want To Convert", font=("jost", 10))
    locationLabel.grid()

    blank = Label()
    blank.grid()

    savebtn = Button(tk, width=20, bg="black", fg="white", text="Choose .docx File", command=openLocation)
    savebtn.grid()

    locationError = Label(tk, text=" ", fg="red", font=("jost", 11))
    locationError.grid()

    blank = Label()
    blank.grid()

    downloadbtn = Button(tk, width=10, bg="black", fg="white", text="Convert", command=c)
    downloadbtn.grid()

    blank = Label()
    blank.grid()

    ydError = Label(tk, text=" ", fg="red", font=("jost", 12))
    ydError.grid()

    blank = Label()
    blank.grid()

    # Developer
    DeveloperLabel = Label(tk, text="Created by", font=("jost", 12))
    DeveloperLabel.grid()

    DeveloperNameLabel = Label(tk, text="Min Khant Soe (HakHak)", font=("jost", 12))
    DeveloperNameLabel.grid()

    tk.mainloop()


if __name__ == '__main__':
    main()
