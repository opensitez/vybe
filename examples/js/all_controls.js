// All Controls Demo — shows every control type from JavaScript
let form = gui.createForm("All Controls");
gui.setProperty(form, "Width", 700);
gui.setProperty(form, "Height", 600);

let y = 10;

// Label
gui.addControl(form, "Label", "lbl1", 10, y, 200, 25);
gui.setProperty("lbl1", "Text", "Label: Hello World");

// Button
y = y + 30;
gui.addControl(form, "Button", "btn1", 10, y, 120, 30);
gui.setProperty("btn1", "Text", "Click Me");

// TextBox
y = y + 35;
gui.addControl(form, "Label", "lblTxt", 10, y, 80, 25);
gui.setProperty("lblTxt", "Text", "TextBox:");
gui.addControl(form, "TextBox", "txt1", 90, y, 200, 25);
gui.setProperty("txt1", "Text", "Type here...");

// CheckBox
y = y + 30;
gui.addControl(form, "CheckBox", "chk1", 10, y, 200, 25);
gui.setProperty("chk1", "Text", "Check me");

// RadioButton
y = y + 30;
gui.addControl(form, "RadioButton", "rad1", 10, y, 120, 25);
gui.setProperty("rad1", "Text", "Option A");
gui.addControl(form, "RadioButton", "rad2", 140, y, 120, 25);
gui.setProperty("rad2", "Text", "Option B");

// ComboBox
y = y + 30;
gui.addControl(form, "Label", "lblCmb", 10, y, 80, 25);
gui.setProperty("lblCmb", "Text", "ComboBox:");
gui.addControl(form, "ComboBox", "cmb1", 90, y, 150, 25);

// NumericUpDown
y = y + 30;
gui.addControl(form, "Label", "lblNum", 10, y, 80, 25);
gui.setProperty("lblNum", "Text", "Numeric:");
gui.addControl(form, "NumericUpDown", "num1", 90, y, 100, 25);

// DateTimePicker
y = y + 30;
gui.addControl(form, "Label", "lblDate", 10, y, 80, 25);
gui.setProperty("lblDate", "Text", "Date:");
gui.addControl(form, "DateTimePicker", "dt1", 90, y, 200, 25);

// ProgressBar
y = y + 30;
gui.addControl(form, "Label", "lblProg", 10, y, 80, 25);
gui.setProperty("lblProg", "Text", "Progress:");
gui.addControl(form, "ProgressBar", "prog1", 90, y, 200, 25);
gui.setProperty("prog1", "Text", "60");

// TrackBar
y = y + 30;
gui.addControl(form, "Label", "lblTrack", 10, y, 80, 25);
gui.setProperty("lblTrack", "Text", "Slider:");
gui.addControl(form, "TrackBar", "track1", 90, y, 200, 30);

// Panel with border
y = y + 40;
gui.addControl(form, "Panel", "panel1", 10, y, 280, 60);
gui.setProperty("panel1", "Text", "Panel (container)");

// LinkLabel
y = y + 70;
gui.addControl(form, "LinkLabel", "link1", 10, y, 200, 25);
gui.setProperty("link1", "Text", "Click this link");

// Right column
let rx = 350;
let ry = 10;

// ListBox
gui.addControl(form, "Label", "lblList", rx, ry, 80, 25);
gui.setProperty("lblList", "Text", "ListBox:");
gui.addControl(form, "ListBox", "list1", rx, ry + 25, 150, 100);

// RichTextBox
ry = ry + 130;
gui.addControl(form, "Label", "lblRich", rx, ry, 100, 25);
gui.setProperty("lblRich", "Text", "RichTextBox:");
gui.addControl(form, "RichTextBox", "rich1", rx, ry + 25, 200, 80);
gui.setProperty("rich1", "Text", "Rich text content here...");

// PictureBox
ry = ry + 110;
gui.addControl(form, "PictureBox", "pic1", rx, ry, 150, 100);

// Status bar at bottom
gui.addControl(form, "StatusStrip", "status1", 0, 570, 700, 25);
gui.setProperty("status1", "Text", "Ready — All controls rendered from JavaScript");

// Event: button click updates status
gui.onEvent("btn1", "Click", () => {
    gui.setProperty("status1", "Text", "Button clicked!");
    gui.setProperty("prog1", "Text", "100");
});

console.log("Launching all controls demo...");
gui.runApplication(form);
