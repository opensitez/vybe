// vybe-test: js/design_patterns/command_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class TextEditor {
    constructor() { this.text = ""; this.history = []; }
    execute(command) {
        this.text = command.execute(this.text);
        this.history.push(command);
    }
    undo() {
        const command = this.history.pop();
        if (command) this.text = command.undo(this.text);
    }
}
const append = text => ({
    execute: current => current + text,
    undo: current => current.slice(0, -text.length)
});
const editor = new TextEditor();
editor.execute(append("Hello"));
editor.execute(append(" World"));
__check(__line(editor.text), "Hello World");
editor.undo();
__check(__line(editor.text), "Hello");
