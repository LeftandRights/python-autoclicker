import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter.font as tkFont

import win32api, win32con
import time, pyautogui
import keyboard, threading

from pynput import mouse
from PIL import Image, ImageTk, ImageGrab


class mainWindow(tk.Tk):
    BACKGROUND_COLOR = "#181820"
    INNERGROUND_COLOR = "#13131a"

    COLUMNS = ("Position", "RGB Target", "Delay after Click", "Wait until Finish", "Move before scan", "Confidence")
    SCRIPT_STATE = 0

    def __init__(self) -> None:
        super().__init__()
        self.title("Windows AutoClicker")

        self.window_width = int((self.winfo_screenwidth()) * 0.56)
        self.window_height = int((self.winfo_screenheight() - 40) * 0.5)
        self.fontSize = int(self.winfo_fpixels("1i") / 11)

        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=self.fontSize)

        self.geometry(f"{self.window_width}x{self.window_height}")

        self.configure(bg=self.BACKGROUND_COLOR)
        # self.resizable(False, False)

        self.option_add("*Font", default_font)

        for objectName in [object for object in dir(self) if object.startswith("property_")]:
            getattr(self, objectName)()

        # for _ in range(1, 6):
        #     self.tableTree.insert("", "end", values=(f"{20*_},{20*_}", "-", "0", "False", "False", "-"))

        keyboard.add_hotkey("ctrl+f2", lambda: threading.Thread(target=self._startScript).start())
        keyboard.add_hotkey("ctrl+alt", lambda: threading.Thread(target=self._add_hotkey).start())
        keyboard.add_hotkey("ctrl+f1", lambda: self._start_recording(True))
        keyboard.add_hotkey(
            "ctrl+1",
            lambda: (
                self.regionTopLeft.config(state="normal"),
                self.regionTopLeft.delete(0, tk.END),
                self.regionTopLeft.insert(0, ", ".join(list(map(str, pyautogui.position())))),
                self.regionTopLeft.config(state="disabled"),
            ),
        )
        keyboard.add_hotkey(
            "ctrl+2",
            lambda: (
                self.regionBottomRight.config(state="normal"),
                self.regionBottomRight.delete(0, tk.END),
                self.regionBottomRight.insert(0, ", ".join(list(map(str, pyautogui.position())))),
                self.regionBottomRight.config(state="disabled"),
            ),
        )

        self.mouseListener = None

    def property_tkinter_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background=self.INNERGROUND_COLOR,
            foreground="white",
            fieldbackground=self.INNERGROUND_COLOR,
            highlightcolor="#00ff3c",
            highlightbackground="#00ff3c",
            highlightthickness=1,
            relief="solid",
        )
        style.configure("Custom.Treeview.Heading", background="#1c1c26", foreground="white", borderwidth=0, font=("Arial", self.fontSize))
        style.layout("Custom.Treeview", [("Custom.Treeview.treearea", {"sticky": "nsew"})])
        style.map("Custom.Treeview", background=[("selected", "#1c1c26")])

    def property_create_table(self) -> None:
        tableFrame = tk.Frame(self, bg=self.BACKGROUND_COLOR, highlightbackground="black", highlightthickness=1, highlightcolor="black")

        # tableFrame.pack(side=tk.RIGHT, anchor="ne")
        tableFrame.place(relwidth=0.526, relheight=0.686, anchor="ne", relx=1.0, rely=0.0)
        tableFrame.pack_propagate(False)

        self.tableTree = ttk.Treeview(tableFrame, columns=self.COLUMNS, show="headings", style="Custom.Treeview")
        self.tableTree.pack(expand=True, fill=tk.BOTH)

        self.tableTree.tag_configure("running", background="#7f818d", foreground="white")

        for column in self.COLUMNS:
            self.tableTree.heading(column, text=column)
            self.tableTree.column(column, anchor="center", width=10)

    def property_imageTab(self) -> None:
        self.currentImageIndex = 1
        self.imageIndex = lambda: f"<Image-{self.currentImageIndex}>"
        self.screenshotList: dict[str, Image.Image] = {}
        self.currentImageData: Image.Image | None = None
        self.isGrayScaleScan = tk.BooleanVar()

        self.imageFrame = tk.Frame(self, bg=self.BACKGROUND_COLOR, highlightbackground="black", highlightthickness=1, highlightcolor="black")
        self.imageFrame.place(relx=0.02, rely=0.49, relwidth=0.42, relheight=0.30)

        self.imageFixedWidth = int(self.window_width * 0.42)
        self.imageFixedHeight = int(self.window_height * 0.30)

        self.imageLabel = tk.Label(self.imageFrame, bg=self.INNERGROUND_COLOR, fg="white", text="Screenshots will be shown here.")
        self.imageLabel.pack(fill=tk.BOTH, expand=True)

        self.pasteScreenshotButton = tk.Button(
            self,
            text="Paste screenshot image",
            bg="#7f818d",
            fg="white",
            highlightthickness=0,
            borderwidth=0,
            command=self._grabClipboard_image,
        )

        self.pasteScreenshotButton.place(relx=0.02, rely=0.804, relwidth=0.20)

        self.deleteLabelScreenshot = tk.Button(
            self,
            text="Remove screenshot",
            bg="#7f818d",
            fg="white",
            highlightthickness=0,
            borderwidth=0,
            command=self._deleteScreenshotsLabel,
            state="disabled",
        )

        self.deleteLabelScreenshot.place(relx=0.239, rely=0.804, relwidth=0.20)

        self.addScreenshotButton = tk.Button(
            self,
            text="Add to table",
            bg="#7f818d",
            fg="white",
            highlightthickness=0,
            borderwidth=0,
            command=self._add_screenshot,
            state="disabled",
        )

        self.addScreenshotButton.place(relx=0.02, rely=0.85, relwidth=0.42)

        tk.Label(self, text="Confidence level (1 - 100):", bg=self.BACKGROUND_COLOR, fg="#fff").place(relx=0.02, rely=0.9)
        self.confidenceLevel = self._create_button(
            master=self, background=self.BACKGROUND_COLOR, fg="#fff", relwidth=0.16, relheight=0.034, relx=0.163, rely=0.902
        )

        self.grayScaleScan = tk.Checkbutton(
            self,
            fg="white",
            borderwidth=0,
            highlightthickness=0,
            text="Grayscale Scan",
            bg=self.BACKGROUND_COLOR,
            variable=self.isGrayScaleScan,
            selectcolor=self.INNERGROUND_COLOR,
            activebackground=self.INNERGROUND_COLOR,
        )

        self.grayScaleScan.place(relx=0.34, rely=0.905, relwidth=0.10)

    def property_advancedTab(self) -> None:
        self.isRegionEnabled = False

        advancedFrame = tk.Frame(self, bg=self.INNERGROUND_COLOR, highlightbackground="black", highlightthickness=1, highlightcolor="black")
        # advancedFrame.place(x=360, y=339)
        advancedFrame.place(
            relx=1.0, rely=1.0, anchor="se", x=0, y=-10, relwidth=0.526, relheight=0.3  # This places it at the bottom-right corner
        )

        self.startButton_text = tk.StringVar(self)
        self.isUnlimitedLoops = tk.BooleanVar()
        self.startButton_text.set("Start (CTRL + F2)")

        self.startButton = tk.Button(
            self,
            textvariable=self.startButton_text,
            bg="#7f818d",
            fg="white",
            highlightthickness=0,
            borderwidth=0,
            command=lambda: threading.Thread(target=self._startScript).start(),
        )

        self.startButton.place(
            relx=1.0,
            rely=1.0,
            anchor="se",  # This places it at the bottom-right corner
            x=4,
            y=0,  # Optional: Adjust with a little offset from the edges if needed
            relwidth=0.529,
        )
        # self.startButton.pack()

        tk.Label(advancedFrame, text="Loop amount", bg=self.INNERGROUND_COLOR, fg="white").place(relx=0.016, rely=0.08)
        self.loopEntry = self._create_button(0.40, 0.12, "#1d1d27", "#fff", master=advancedFrame, relx=0.02, rely=0.2)

        self.unlimitedLoop = tk.Checkbutton(
            advancedFrame,
            text="Unlimited loops",
            bg=self.INNERGROUND_COLOR,
            fg="white",
            highlightthickness=0,
            borderwidth=0,
            selectcolor=self.INNERGROUND_COLOR,
            activebackground=self.INNERGROUND_COLOR,
            variable=self.isUnlimitedLoops,
            onvalue=True,
            offvalue=False,
            command=lambda: self.loopEntry.config(state=("disabled" if self.isUnlimitedLoops.get() else "normal")),
        )

        self.unlimitedLoop.place(relx=0.016, rely=0.35)

        # tk.Label(advancedFrame, text="Region (Area of the screen to be scan)", bg=self.INNERGROUND_COLOR, fg="white").place(
        #     relx=0.012, rely=0.55
        # )
        # self.region = self._create_button(0.40, 0.12, "#1d1d27", "#fff", master=advancedFrame, relx=0.016, rely=0.68)
        # self.region.config(state="disabled")

        tk.Label(advancedFrame, text="Top-Left Region (Ctrl + 1)", bg=self.INNERGROUND_COLOR, fg="#fff").place(relx=0.55, rely=0.08)
        self.regionTopLeft = self._create_button(0.40, 0.12, "#1d1d27", "#fff", master=advancedFrame, relx=0.55, rely=0.2)
        self.regionTopLeft.config(state="disabled")

        tk.Label(advancedFrame, text="Bottom-Right Region (Ctrl + 2)", bg=self.INNERGROUND_COLOR, fg="#fff").place(relx=0.55, rely=0.38)
        self.regionBottomRight = self._create_button(0.40, 0.12, "#1d1d27", "#fff", master=advancedFrame, relx=0.55, rely=0.53)
        self.regionBottomRight.config(state="disabled")

        self.triggerRegionButton = tk.Button(
            advancedFrame,
            text="Enable Region",
            bg="#7f818d",
            fg="white",
            width=45,
            highlightthickness=0,
            borderwidth=0,
            command=self._regionSwitch,
        )

        self.triggerRegionButton.place(relx=0.55, rely=0.7, relwidth=0.40)

    def property_mainMenu(self) -> None:
        self.waitForFinish = tk.IntVar(self)
        self.delayBeforeAdding = tk.IntVar(self)
        self.moveBeforeScan = tk.BooleanVar(self)

        self.menuFrame = tk.Frame(
            self,
            background=self.INNERGROUND_COLOR,
            highlightthickness=1,
            highlightcolor="black",
            highlightbackground="black",
        )

        self.menuFrame.place(relx=0.02, rely=0.035, relwidth=0.42, relheight=0.30)  # 1 = 0.02
        self.x_axis_text = tk.Label(self.menuFrame, text="X Axis Coordinate", bg=self.INNERGROUND_COLOR, highlightthickness=0, fg="#fff")
        self.x_axis_text.place(relx=0.04, rely=0.06, anchor="nw")

        self.x_axis_Entry = self._create_button(0.42, 0.11, background="#1d1d27", fg="#fff", master=self.menuFrame, rely=0.19, relx=0.04)

        self.y_axis_text = tk.Label(self.menuFrame, text="Y Axis Coordinate", bg=self.INNERGROUND_COLOR, highlightthickness=0, fg="#fff")
        self.y_axis_text.place(relx=0.525, rely=0.06)

        self.y_axis_Entry = self._create_button(
            0.42, 0.11, background="#1d1d27", fg="#fff", master=self.menuFrame, anchor="ne", rely=0.19, relx=0.95
        )

        self.rgb_target_text = tk.Label(
            self.menuFrame, text="RGB Target (hex / rgb)", bg=self.INNERGROUND_COLOR, highlightthickness=0, fg="#fff"
        )
        self.rgb_target_text.place(rely=0.34, relx=0.04)

        self.rgb_Entry = self._create_button(0.42, 0.11, background="#1d1d27", fg="#fff", master=self.menuFrame, rely=0.49, relx=0.04)

        self.delayAfterClick_text = tk.Label(
            self.menuFrame, text="Delay after click", bg=self.INNERGROUND_COLOR, highlightthickness=0, fg="#fff"
        )

        self.delayAfterClick_text.place(rely=0.34, relx=0.525)

        self.delayAfterClick_Entry = self._create_button(
            0.42, 0.11, background="#1d1d27", fg="#fff", master=self.menuFrame, relx=0.95, rely=0.49, anchor="ne"
        )

        self.addButton = tk.Button(
            self.menuFrame,
            text="Add",
            bg="#7f818d",
            fg="white",
            width=45,
            highlightthickness=0,
            borderwidth=0,
            command=self._add_button,
        )
        self.addButton.place(relwidth=0.42, relheight=0.12, relx=0.46, rely=0.67, anchor="ne")

        self.addDirectlyButton = tk.Button(
            self.menuFrame, text="Add directly (Ctrl + Alt)", bg="#7f818d", fg="white", width=45, highlightthickness=0, borderwidth=0
        )
        self.addDirectlyButton.place(relwidth=0.42, relheight=0.12, relx=0.46, rely=0.83, anchor="ne")

        self.waitForFinishButton = tk.Checkbutton(
            self.menuFrame,
            fg="white",
            borderwidth=0,
            highlightthickness=0,
            text="Wait for RGB / Image Scan",
            bg=self.INNERGROUND_COLOR,
            variable=self.waitForFinish,
            selectcolor=self.INNERGROUND_COLOR,
            activebackground=self.INNERGROUND_COLOR,
            command=lambda: self.moveBeforeScanButton.config(state=("disabled" if not self.waitForFinish.get() else "normal")),
            onvalue=1,
            offvalue=0,
        )

        self.waitForFinishButton.place(relx=0.53, rely=0.65)

        self.delayBeforeAddingButton = tk.Checkbutton(
            self.menuFrame,
            fg="white",
            borderwidth=0,
            highlightthickness=0,
            text="1.5s delay before direct adding",
            bg=self.INNERGROUND_COLOR,
            variable=self.delayBeforeAdding,
            selectcolor=self.INNERGROUND_COLOR,
            activebackground=self.INNERGROUND_COLOR,
        )

        self.delayBeforeAddingButton.place(relx=0.53, rely=0.89)

        self.moveBeforeScanButton = tk.Checkbutton(
            self.menuFrame,
            fg="white",
            borderwidth=0,
            highlightthickness=0,
            text="Move to position before scan",
            bg=self.INNERGROUND_COLOR,
            variable=self.moveBeforeScan,
            selectcolor=self.INNERGROUND_COLOR,
            activebackground=self.INNERGROUND_COLOR,
            state="disabled",
            onvalue=True,
            offvalue=False,
        )

        self.moveBeforeScanButton.place(relx=0.53, rely=0.77)

        self.recordButton = tk.Button(
            self,
            text="Record (Ctrl + F1)",
            bg="#7f818d",
            fg="white",
            width=19,
            highlightthickness=0,
            borderwidth=0,
            command=self._start_recording,
        )
        self.recordButton.place(relx=0.02, rely=0.35, relwidth=0.42)

        self.editButton = tk.Button(self, text="Edit", bg="#7f818d", fg="white", highlightthickness=0, borderwidth=0, command=self._edit_button)
        self.editButton.place(relwidth=0.20, rely=0.40, relx=0.02)

        self.removeButton = tk.Button(
            self, text="Delete", bg="#7f818d", fg="white", width=21, highlightthickness=0, borderwidth=0, command=self._delete_button
        )
        self.removeButton.place(relwidth=0.20, rely=0.40, relx=0.24)

    def _regionSwitch(self) -> None:
        if not self.isRegionEnabled:
            if not (self.regionTopLeft.get() and self.regionBottomRight.get()):
                messagebox.showerror("Error", "Please provide both top-left and bottom-right coordinates.")
                return

            self.isRegionEnabled = True
            self.triggerRegionButton.config(text="Disable Region")
            return

        self.isRegionEnabled = False
        self.triggerRegionButton.config(text="Enable Region")

    def _create_button(self, relwidth: float, relheight: float, background: str, fg: str, master=None, **kwargs) -> tk.Entry:
        border_frame = tk.Frame((self if master is None else master), background="#7f818d", highlightthickness=0)
        border_frame.place(relwidth=relwidth, relheight=relheight, **kwargs)

        entry = tk.Entry(
            border_frame,
            background=background,
            fg=fg,
            borderwidth=0,
            highlightthickness=0,
            justify="center",
            disabledbackground="#1d1d27",
        )
        # entry.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        entry.pack(fill=tk.BOTH, expand=True, pady=1, padx=1)

        return entry

    def _deleteScreenshotsLabel(self) -> None:
        if self.currentImageData is None:
            return

        self.imageLabel.config(image=None)
        self.imageLabel.image = None

        self.deleteLabelScreenshot.config(state="disabled")
        self.addScreenshotButton.config(state="disabled")
        self.currentImageData = None

    def _grabClipboard_image(self) -> None:
        if not (image := ImageGrab.grabclipboard()):
            messagebox.showerror("Error", "No image found in clipboard.")
            return

        self.currentImageData = image
        resizedImage = image.resize((self.imageFixedWidth, self.imageFixedHeight), Image.LANCZOS)
        imgTk = ImageTk.PhotoImage(resizedImage)

        self.imageLabel.config(image=imgTk)
        self.imageLabel.image = imgTk

        self.addScreenshotButton.config(state="normal")
        self.deleteLabelScreenshot.config(state="normal")

    def _add_button(self) -> None:
        xPos, yPos, delay, waitToFinish = (
            self.x_axis_Entry.get(),
            self.y_axis_Entry.get(),
            self.delayAfterClick_Entry.get(),
            self.waitForFinish.get() == 1,
        )

        if (data := self._validate_entry()) is not None:
            self.tableTree.insert(
                "",
                "end",
                values=(
                    f"{xPos}, {yPos}",
                    ", ".join([str(_) for _ in data[1]]) if data[1] != "-" else "-",
                    delay,
                    waitToFinish,
                    self.moveBeforeScan.get() == 1,
                    "-",
                ),
            )

    def _add_screenshot(self) -> None:
        if not ((confidenceLevel := self.confidenceLevel.get()).isdigit() and int(confidenceLevel) > 0 and int(confidenceLevel) < 101):
            messagebox.showerror("Error", "Invalid confidence level. Please enter a number between 1 and 100.")
            return

        confidenceLevel = int(confidenceLevel) / 100
        self.screenshotList[self.imageIndex()] = self.currentImageData
        delay = self.delayAfterClick_Entry.get()

        self.tableTree.insert(
            "",
            tk.END,
            values=(
                self.imageIndex(),
                "Grayscale" if self.isGrayScaleScan.get() else "-",
                0 if not delay else delay,
                self.waitForFinish.get() == 1,
                self.moveBeforeScan.get() == 1,
                confidenceLevel,
            ),
        )

        self.currentImageIndex += 1

    def _delete_button(self) -> None:
        if selectedItem := self.tableTree.selection():
            self.tableTree.delete(selectedItem[0])

    def _edit_button(self) -> None:
        if not (selectedItem := self.tableTree.selection()):
            return

        itemData = self.tableTree.item(selectedItem[0])["values"]
        editStatus = False

        def cancelEdit() -> None:
            self.recordButton.config(state="normal")
            self.addButton.config(state="normal")
            self.removeButton.config(text="Delete", command=self._delete_button)
            self.editButton.config(text="Edit", command=self._edit_button)
            self.pasteScreenshotButton.config(state="normal")
            self.imageLabel.config(image=None)
            self.imageLabel.image = None
            self._flush_entry()

        def submitEdit() -> None:
            nonlocal editStatus

            if (not (data := self._validate_entry())) and not itemData[0].startswith("<"):
                print("Stuck")
                return

            xPos, yPos, delay, waitToFinish = (
                self.x_axis_Entry.get(),
                self.y_axis_Entry.get(),
                self.delayAfterClick_Entry.get(),
                self.waitForFinish.get() == 1,
            )

            if not (itemData[0].startswith("<")):
                self.tableTree.item(
                    selectedItem[0],
                    values=(
                        f"{xPos.replace(' ', '')}, {yPos.replace(' ', '')}",
                        "-" if data[1] == "-" else ", ".join(data[1]),
                        delay,
                        waitToFinish,
                        (self.moveBeforeScan.get() == 1) and waitToFinish,
                        "-",
                    ),
                )

            else:
                if not ((confidenceLevel := self.confidenceLevel.get()).isdigit() and int(confidenceLevel) > 0 and int(confidenceLevel) < 101):
                    messagebox.showerror("Error", "Invalid confidence level. Please enter a number between 1 and 100.")
                    editStatus = False
                    print("Failed the confidence level")
                    return

                self.screenshotList[itemData[0]] = self.currentImageData

                self.tableTree.item(
                    selectedItem[0],
                    values=(
                        itemData[0],
                        "Grayscale" if self.isGrayScaleScan.get() else "-",
                        delay,
                        waitToFinish,
                        (self.moveBeforeScan.get() == 1) and waitToFinish,
                        float(confidenceLevel) / 100,
                    ),
                )

            editStatus = True

            if editStatus:
                self.recordButton.config(state="normal")
                self.addButton.config(state="normal")
                self.removeButton.config(text="Delete", command=self._delete_button)
                self.editButton.config(text="Edit", command=self._edit_button)
                self.imageLabel.config(image=None)
                self.imageLabel.image = None
                self._flush_entry()

        self._flush_entry()

        if not (itemData[0].startswith("<")):
            self.x_axis_Entry.insert(0, itemData[0].split(",")[0])
            self.y_axis_Entry.insert(0, itemData[0].split(",")[1])
        else:
            image = self.screenshotList[itemData[0]]
            resizedImage = image.resize((self.imageFixedWidth, self.imageFixedHeight), Image.LANCZOS)
            imgTk = ImageTk.PhotoImage(resizedImage)

            self.imageLabel.config(image=imgTk)
            self.imageLabel.image = imgTk
            self.currentImageData = image

        if itemData[1] != "Grayscale":
            self.rgb_Entry.insert(0, itemData[1].replace("(", "").replace(")", "").replace("-", ""))
        else:
            self.isGrayScaleScan.set(True)

        self.delayAfterClick_Entry.insert(0, itemData[2])
        self.waitForFinish.set(_ := 1 if itemData[3] == "True" else 0)
        self.moveBeforeScan.set(1 if itemData[4] == "True" else 0)

        if itemData[5] != "-":
            self.confidenceLevel.insert(0, int(float(itemData[5]) * 100))

        if _:
            self.moveBeforeScanButton.config(state="normal")

        self.recordButton.config(state="disabled")
        self.addButton.config(state="disabled")
        self.removeButton.config(text="Submit Edit", command=submitEdit)
        self.editButton.config(text="Cancel Edit", command=cancelEdit)
        self.deleteLabelScreenshot.config(state="disabled")

    def _flush_entry(self) -> None:
        self.x_axis_Entry.delete(0, tk.END)
        self.y_axis_Entry.delete(0, tk.END)
        self.rgb_Entry.delete(0, tk.END)
        self.confidenceLevel.delete(0, tk.END)
        self.isGrayScaleScan.set(False)

        self.delayAfterClick_Entry.delete(0, tk.END)
        self.waitForFinish.set(0)
        self.moveBeforeScan.set(0)
        self.moveBeforeScanButton.config(state="disabled")

    def _validate_entry(self) -> tuple[bool, list | None]:
        xPos, yPos, rgbTarget, delay = (
            self.x_axis_Entry.get(),
            self.y_axis_Entry.get(),
            self.rgb_Entry.get(),
            self.delayAfterClick_Entry.get(),
        )

        if not all([xPos, yPos, delay]):
            return

        try:
            int(xPos), int(yPos)
            float(delay)

        except ValueError:
            messagebox.showerror("Error", "Invalid input")

        rgb = "-"
        _err = lambda: messagebox.showerror("Error", "Invalid RGB Target, must be in hex or RGB format, e.g #ffffff or 255,255,255")

        if rgbTarget:
            if not rgbTarget.startswith("#"):
                rgb = tuple(color for color in rgbTarget.replace(" ", "").split(","))

                if (len(rgb) != 3) or not all(color.isdigit() for color in rgb):
                    _err()
                    return

            else:
                if len(rgbTarget) != 7:
                    _err()
                    return

                try:
                    rgbTarget = rgbTarget.lstrip("#")
                    rgb = tuple(int(rgbTarget[i : i + 2], 16) for i in (0, 2, 4))
                except Exception:
                    _err()
                    return

        return True, rgb

    def _startScript(self, *args) -> None:
        def script() -> None:
            for item_id in self.tableTree.get_children():
                if self.SCRIPT_STATE == 0:
                    break

                self.tableTree.selection_set(item_id)
                self.tableTree.focus(item_id)

                self.tableTree.item(item_id, tags=("running",))
                position, rgb_target, delay, waitToFinish, moveBeforeScan, confidenceLevel = self.tableTree.item(item_id, "values")

                if rgb_target != "Grayscale":
                    rgb_target = tuple(int(i) for i in rgb_target.split(",")) if len(rgb_target) > 1 else "-"

                if not position.startswith("<"):
                    x, y = map(int, position.split(","))

                delay = float(delay)
                waitToFinish = True if waitToFinish == "True" else False

                if isinstance(rgb_target, tuple):

                    if moveBeforeScan == "True":
                        win32api.SetCursorPos((x, y))

                    while self.SCRIPT_STATE == 1 and waitToFinish:
                        if pyautogui.pixelMatchesColor(x, y, rgb_target) and self.SCRIPT_STATE == 1:
                            break

                        time.sleep(1)

                    if not pyautogui.pixelMatchesColor(x, y, rgb_target):
                        continue

                elif position.startswith("<"):
                    kwargs = {}

                    if self.isRegionEnabled:
                        topRightRegion = list(map(int, self.regionTopLeft.get().split(", ")))
                        bottomLeftRegion = list(map(int, self.regionBottomRight.get().split(", ")))

                        width = bottomLeftRegion[0] - topRightRegion[0]
                        height = bottomLeftRegion[1] - topRightRegion[1]

                        kwargs["region"] = (topRightRegion[0], topRightRegion[1], width, height)

                    location = None

                    while True:
                        try:
                            location = pyautogui.locateOnScreen(
                                self.screenshotList[position],
                                grayscale=(True if rgb_target == "Grayscale" else False),
                                confidence=float(confidenceLevel),
                                **kwargs,
                            )

                        except (pyautogui.ImageNotFoundException, OSError):
                            if waitToFinish:
                                time.sleep(0.5)
                                continue

                        break

                    if location == None:
                        continue

                    x, y = location.left + (location.width // 2), location.top + (location.height // 2)

                win32api.SetCursorPos((x, y))
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                time.sleep(delay)

                self.tableTree.item(item_id, tags=())

        if self.SCRIPT_STATE == 1:
            self.startButton_text.set("Start (CTRL + F2)")
            self.SCRIPT_STATE = 0
            return

        self.startButton_text.set("Stop (CTRL + F2)")
        self.SCRIPT_STATE = 1

        if self.isUnlimitedLoops.get():
            while self.SCRIPT_STATE == 1:
                script()

        else:
            try:
                for _ in range(int(self.loopEntry.get())):
                    script()

            except ValueError:
                messagebox.showerror("Error", "Invalid loop count")

        self.startButton_text.set("Start (CTRL + F2)")
        self.SCRIPT_STATE = 0

    def _on_mouse_click(self, x, y, button, pressed) -> None:
        if pressed and button == mouse.Button.left:
            self.tableTree.insert(
                "",
                tk.END,
                values=(f"{x},{y}", ", ".join(str(_) for _ in pyautogui.pixel(x, y)), "0.5", "False", "False", "-"),
            )

    def _start_recording(self, triggeredByHotkey: bool = False) -> None:
        if self.mouseListener is None:
            self.mouseListener = mouse.Listener(on_click=self._on_mouse_click)
            self.recordButton.config(text="Stop Recording (CTRL + F1)")
            self.mouseListener.start()
            return

        if not triggeredByHotkey:
            self.tableTree.delete(self.tableTree.get_children()[-1])

        self.recordButton.config(text="Record (CTRL + F1)")
        self.mouseListener.stop()
        self.mouseListener = None

    def _add_hotkey(self) -> None:
        x, y = pyautogui.position()

        if self.delayBeforeAdding.get() == 1:
            time.sleep(1.5)

        self.tableTree.insert(
            "",
            tk.END,
            values=(
                f"{x},{y}",
                ", ".join([str(_) for _ in pyautogui.pixel(x, y)]),
                "0.5",
                _ := self.waitForFinish.get() == 1,
                self.moveBeforeScan.get() and _,
                "-",
            ),
        )

        self.update_idletasks()
        self.update()


mainWindow().mainloop()
