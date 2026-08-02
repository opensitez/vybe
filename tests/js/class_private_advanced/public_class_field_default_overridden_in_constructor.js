// vybe-test: js/class_private_advanced/public_class_field_default_overridden_in_constructor
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class Widget {
    color = "grey";
    constructor(color) {
        if (color) this.color = color;
    }
}
const w1 = new Widget("blue");
const w2 = new Widget();
__check(__line(w1.color), "blue");
__check(__line(w2.color), "grey");
