// vybe-test: js/number_edge_cases/float_precision_epsilon
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

const a = 0.1 + 0.2;
__check(__line(a === 0.3), "false");
__check(__line(Math.abs(a - 0.3) < Number.EPSILON), "true");
