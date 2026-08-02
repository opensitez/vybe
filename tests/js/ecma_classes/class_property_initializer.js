// vybe-test: js/ecma_classes/class_property_initializer
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

class Defaults {
    name = "unnamed";
    count = 0;
    items = [];

    describe() {
        return this.name + ":" + this.count;
    }
}
const d = new Defaults();
__check(__line(d.describe()), "unnamed:0");
