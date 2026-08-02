// vybe-test: js/spread_rest_advanced/rest_and_spread_roundtrip
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

function pass(...args) { return args; }
const src = [1, 2, 3];
const result = pass(...src);
__check(__line(result.join(",")), "1,2,3");
