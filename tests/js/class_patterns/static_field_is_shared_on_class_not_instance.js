// vybe-test: js/class_patterns/static_field_is_shared_on_class_not_instance
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

class Box {
    static count = 2;
}
let b = new Box();
__check(__line(Box.count), "2");
__check(__line(b.count), "undefined");
