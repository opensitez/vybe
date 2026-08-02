// vybe-test: js/ecma_classes/class_has_private_field_with_getter_and_setter
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

class Box {
    #value = 1;
    get value() { return this.#value; }
    set value(v) { this.#value = v; }
}
const b = new Box();
b.value = 5;
__check(__line(b.value), "5");
