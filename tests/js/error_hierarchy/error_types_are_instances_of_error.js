// vybe-test: js/error_hierarchy/error_types_are_instances_of_error
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

const errors = [
    new TypeError("t"),
    new RangeError("r"),
    new ReferenceError("ref"),
    new SyntaxError("s"),
    new URIError("u"),
    new EvalError("e"),
];
__check(__line(errors.every(e => e instanceof Error)), "true");
__check(__line(errors.map(e => e.constructor.name).join(",")), "TypeError,RangeError,ReferenceError,SyntaxError,URIError,EvalError");
