// vybe-test: js/class_patterns/valueof_override
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

class Num {
    constructor(v) { this.v = v; }
    valueOf() { return this.v; }
}
let a = new Num(10);
let b = new Num(20);
__check(__line(a + b), "30");
__check(__line(a * 3), "30");
