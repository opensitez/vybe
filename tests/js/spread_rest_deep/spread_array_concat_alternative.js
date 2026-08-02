// vybe-test: js/spread_rest_deep/spread_array_concat_alternative
// origin: languages/js/tests/js/test_spread_rest_deep.rs

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

const a = [1, 2, 3];
const b = [4, 5, 6];
const combined = [...a, ...b];
__check(__line(combined.join(",")), "1,2,3,4,5,6");
