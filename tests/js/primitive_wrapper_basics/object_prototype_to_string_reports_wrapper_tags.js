// vybe-test: js/primitive_wrapper_basics/object_prototype_to_string_reports_wrapper_tags
// origin: languages/js/tests/js/test_primitive_wrapper_basics.rs

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

__check(__line(Object.prototype.toString.call(new Number(1))), "[object Number]");
__check(__line(Object.prototype.toString.call(new String("x"))), "[object String]");
__check(__line(Object.prototype.toString.call(new Boolean(false))), "[object Boolean]");
