// Every control type, as the HTML elements they actually are.
//
// The WinForms names have HTML equivalents and this is the whole mapping:
// Label→<label>, Button→<button>, TextBox→<input>, CheckBox/RadioButton→
// <input type>, ComboBox→<select>, NumericUpDown→<input type=number>,
// DateTimePicker→<input type=date>, ProgressBar→<progress>, TrackBar→
// <input type=range>, Panel→<div>, LinkLabel→<a>, ListBox→<select multiple>,
// RichTextBox→<textarea>, PictureBox→<img>, StatusStrip→a <div> at the end.
//
// Nothing here places a control at an x/y. Elements are appended and the
// document lays them out, which is the point of building on the web platform.
document.setTitle("All Controls");

function el(tag, type) {
    let node = type === undefined
        ? document.createElement(tag)
        : document.createElement(tag, type);
    document.body.appendChild(node);
    return node;
}

function labelled(text, tag, type) {
    let row = el("div");
    let caption = document.createElement("label");
    caption.appendChild(document.createTextNode(text));
    row.appendChild(caption);
    let field = type === undefined
        ? document.createElement(tag)
        : document.createElement(tag, type);
    row.appendChild(field);
    return field;
}

let label = el("label");
label.appendChild(document.createTextNode("Label: Hello World"));

let button = el("button");
button.appendChild(document.createTextNode("Click Me"));

let text = labelled("TextBox:", "input");
text.setAttribute("value", "Type here...");

let check = labelled("Check me", "input", "checkbox");

let optionA = labelled("Option A", "input", "radio");
optionA.setAttribute("name", "choice");
let optionB = labelled("Option B", "input", "radio");
optionB.setAttribute("name", "choice");

let combo = labelled("ComboBox:", "select");

let numeric = labelled("Numeric:", "input", "number");

let date = labelled("Date:", "input", "date");

let progress = labelled("Progress:", "progress");
progress.setAttribute("value", "60");
progress.setAttribute("max", "100");

let slider = labelled("Slider:", "input", "range");

let panel = el("div");
panel.appendChild(document.createTextNode("Panel (container)"));

let link = el("a");
link.setAttribute("href", "#");
link.appendChild(document.createTextNode("Click this link"));

let list = labelled("ListBox:", "select");
list.setAttribute("multiple", "");

let rich = labelled("RichTextBox:", "textarea");
rich.setAttribute("value", "Rich text content here...");

let picture = el("img");

let status = el("div");
status.appendChild(document.createTextNode("Ready — every control is an element"));

button.addEventListener("click", () => {
    status.setTextContent("Button clicked!");
    progress.setAttribute("value", "100");
});

console.log("Built every control as a DOM element.");
