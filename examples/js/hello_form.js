// Hello Form — the same UI, built the way a browser takes it. There is no
// toolkit API and no `gui` namespace: a control IS a DOM element, so this is
// `createElement` / `appendChild` / `addEventListener` against the document the
// page already has. `document` is a global the web platform binds, exactly as
// `window.document` is in a browser — nothing to import.
document.setTitle("Hello from JS!");

function add(tag, text) {
    let node = document.createElement(tag);
    node.setTextContent(text);
    document.body.appendChild(node);
    return node;
}

add("div", "Hello World from JavaScript!");

let counterLabel = add("div", "Click count: 0");

let clickCount = 0;
let button = add("button", "Click Me!");
button.addEventListener("click", () => {
    clickCount = clickCount + 1;
    counterLabel.setTextContent("Click count: " + clickCount);
    button.setTextContent("Clicked " + clickCount + "x");
});

let reset = add("button", "Reset");
reset.addEventListener("click", () => {
    clickCount = 0;
    counterLabel.setTextContent("Click count: 0");
    button.setTextContent("Click Me!");
});

let input = document.createElement("input");
input.setAttribute("value", "Type here...");
document.body.appendChild(input);

add(
    "div",
    "This form was created entirely from JavaScript, running on the Vybe bytecode VM against the WHATWG DOM.",
);

// Nothing launches the page. It runs because it HAS a document with content.
console.log("Built the page from JavaScript.");
