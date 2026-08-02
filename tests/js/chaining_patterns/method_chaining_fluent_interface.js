// vybe-test: js/chaining_patterns/method_chaining_fluent_interface
// origin: languages/js/tests/js/test_chaining_patterns.rs

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

class StringBuilder {
    #parts = [];
    append(str) { this.#parts.push(str); return this; }
    prepend(str) { this.#parts.unshift(str); return this; }
    join(sep = "") { return this.#parts.join(sep); }
    toString() { return this.join(); }
}
const result = new StringBuilder()
    .append("World")
    .prepend("Hello ")
    .append("!")
    .toString();
__check(__line(result), "Hello World!");
