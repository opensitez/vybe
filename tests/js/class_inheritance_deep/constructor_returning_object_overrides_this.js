// vybe-test: js/class_inheritance_deep/constructor_returning_object_overrides_this
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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

class Weird {
    constructor() {
        return { custom: true };
    }
}
const w = new Weird();
__check(__line(w.custom), "true");
__check(__line(w instanceof Weird), "false");
