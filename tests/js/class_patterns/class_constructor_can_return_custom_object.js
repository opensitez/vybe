// vybe-test: js/class_patterns/class_constructor_can_return_custom_object
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

class Weird {
    constructor() {
        this.value = 1;
        return { value: 99 };
    }
}
let w = new Weird();
__check(__line(w.value), "99");
