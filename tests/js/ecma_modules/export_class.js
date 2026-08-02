// vybe-test: js/ecma_modules/export_class
// origin: languages/js/tests/js/test_ecma_modules.rs

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

export class MyClass {
    constructor(x) { this.x = x; }
    get() { return this.x; }
}
const m = new MyClass(42);
__check(__line(m.get()), "42");
