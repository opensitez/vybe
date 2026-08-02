// vybe-test: js/operator_misc/in_operator_with_symbol_and_non_object_rhs_throws
// origin: languages/js/tests/js/test_operator_misc.rs

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

const key = Symbol("x");
try {
    console.log(key in 42);
} catch (e) {
    console.log(e.name);
}
