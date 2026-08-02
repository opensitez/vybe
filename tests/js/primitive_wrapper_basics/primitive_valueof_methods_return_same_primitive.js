// vybe-test: js/primitive_wrapper_basics/primitive_valueof_methods_return_same_primitive
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

__check(__line("hi".valueOf()), "hi");
__check(__line((3.14).valueOf()), "3.14");
__check(__line(true.valueOf()), "true");
