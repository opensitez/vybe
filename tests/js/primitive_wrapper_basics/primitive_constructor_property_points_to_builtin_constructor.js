// vybe-test: js/primitive_wrapper_basics/primitive_constructor_property_points_to_builtin_constructor
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

__check(__line("hi".constructor === String), "true");
__check(__line((42).constructor === Number), "true");
__check(__line(true.constructor === Boolean), "true");
