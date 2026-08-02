// vybe-test: js/ecma_operators/in_operator_on_array_indexes
// origin: languages/js/tests/js/test_ecma_operators.rs

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

const arr = ["x", "y"];
__check(__line(0 in arr), "true");
__check(__line(2 in arr), "false");
__check(__line("length" in arr), "true");
