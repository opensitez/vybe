// vybe-test: js/oop_patterns_advanced/memento_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class Editor {
    #content = "";
    #history = [];
    type(text) { this.#history.push(this.#content); this.#content += text; }
    undo() { if (this.#history.length) this.#content = this.#history.pop(); }
    get content() { return this.#content; }
}
const e = new Editor();
e.type("Hello");
e.type(" World");
__check(__line(e.content), "Hello World");
e.undo();
__check(__line(e.content), "Hello");
e.undo();
__check(__line(e.content), "");
