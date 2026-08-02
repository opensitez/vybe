// vybe-test: js/number_edge_cases/max_safe_integer_boundary
// origin: languages/js/tests/js/test_number_edge_cases.rs

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

const max = Number.MAX_SAFE_INTEGER;
__check(__line(max), "9007199254740991");
__check(__line(Number.isSafeInteger(max)), "true");
__check(__line(Number.isSafeInteger(max + 1)), "false");
__check(__line(max + 1 === max + 2), "true"); // precision loss
