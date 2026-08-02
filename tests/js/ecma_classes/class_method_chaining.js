// vybe-test: js/ecma_classes/class_method_chaining
// origin: languages/js/tests/js/test_ecma_classes.rs

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

class Builder {
    constructor() { this.parts = []; }
    add(part) { this.parts.push(part); return this; }
    build() { return this.parts.join(", "); }
}
const result = new Builder().add("a").add("b").add("c").build();
__check(__line(result), "a, b, c");
