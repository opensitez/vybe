// vybe-test: js/class_private_deep/private_method_calling_private_method
// origin: languages/js/tests/js/test_class_private_deep.rs

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

class Parser {
    #tokenize(str) { return str.split(" "); }
    #process(tokens) { return tokens.map(t => t.toUpperCase()); }
    parse(str) { return this.#process(this.#tokenize(str)).join(","); }
}
const p = new Parser();
__check(__line(p.parse("hello world foo")), "HELLO,WORLD,FOO");
