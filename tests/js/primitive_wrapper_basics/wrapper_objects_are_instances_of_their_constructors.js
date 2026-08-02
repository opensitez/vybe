// vybe-test: js/primitive_wrapper_basics/wrapper_objects_are_instances_of_their_constructors
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

__check(__line(new Number(1) instanceof Number), "true");
__check(__line(new String("x") instanceof String), "true");
__check(__line(new Boolean(true) instanceof Boolean), "true");
