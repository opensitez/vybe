// vybe-test: js/spread_rest_deep/spread_converts_string_to_chars
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

const chars = [..."hello"];
__check(__line(chars.join("-")), "h-e-l-l-o");
