// vybe-test: js/spread_rest_advanced/object_rest_empty_when_all_named
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

const { x, y, ...rest } = { x: 1, y: 2 };
__check(__line(Object.keys(rest).length), "0");
