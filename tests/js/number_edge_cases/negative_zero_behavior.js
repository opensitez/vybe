// vybe-test: js/number_edge_cases/negative_zero_behavior
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

const negZero = -0;
__check(__line(negZero === 0), "true");
__check(__line(Object.is(negZero, 0)), "false");
__check(__line(Object.is(negZero, -0)), "true");
__check(__line(String(negZero)), "0");
__check(__line(1 / negZero), "-Infinity");
