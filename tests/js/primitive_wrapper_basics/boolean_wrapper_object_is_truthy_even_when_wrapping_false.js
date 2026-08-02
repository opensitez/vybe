// vybe-test: js/primitive_wrapper_basics/boolean_wrapper_object_is_truthy_even_when_wrapping_false
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

if (new Boolean(false)) {
  console.log("truthy");
} else {
  console.log("falsey");
}
