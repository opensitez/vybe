// vybe-test: js/coercion_toprimitive/valueof_in_arithmetic
// origin: languages/js/tests/js/test_coercion_toprimitive.rs

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

const obj = { valueOf() { return 42; } };
__check(__line(obj + 8), "50");
__check(__line(obj * 2), "84");
__check(__line(obj - 10), "32");
