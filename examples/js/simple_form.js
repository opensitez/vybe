// The same form, written the way a browser would take it: no toolkit API and
// no `gui` namespace — a control IS a DOM element, so this is `createElement`,
// `appendChild`, and `addEventListener` against the document the page has.
//
// Nothing tells the page to run. It runs because it HAS a document.
document.setTitle("Test");

let label = document.createElement("div");
label.setTextContent("Not clicked yet");
document.body.appendChild(label);

let button = document.createElement("button");
button.setTextContent("Click");
document.body.appendChild(button);

button.addEventListener("click", () => {
    label.setTextContent("Clicked!");
});
