// vybe-test: js/class_patterns/symbol_to_primitive_default_hint_used_in_addition
// origin: languages/js/tests/js/test_class_patterns.rs

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

class Amount {
    constructor(v) { this.v = v; }
    [Symbol.toPrimitive](hint) {
        __check(__line(hint), "default");
        return this.v;
    }
}
let a = new Amount(7);
__check(__line(a + 5), "12");
