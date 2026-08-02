// vybe-test: js/class_fields_advanced/arrow_field_vs_method_binding
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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

class Handler {
    name = "handler";
    // Arrow as field binds 'this' permanently
    arrowMethod = () => this.name;
    // Regular method — 'this' depends on call site
    regularMethod() { return this.name; }
}
const h = new Handler();
const { arrowMethod, regularMethod } = h;
__check(__line(arrowMethod()), "handler"); // works — bound
let threw = false;
try {
    const r = regularMethod(); // might throw or return undefined
} catch { threw = true; }
// Either throws (strict mode) or returns undefined (sloppy)
__check(__line(typeof arrowMethod()), "string");
