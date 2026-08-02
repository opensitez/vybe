// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_builtin_array_subclassing
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class CustomArray extends Array {
    first() { return this[0]; }
    last() { return this[this.length - 1]; }
}
const arr = new CustomArray(10, 20, 30);
__check(__line(arr.first() + "|" + arr.last() + "|" + (arr instanceof Array)), "10|30|true");
