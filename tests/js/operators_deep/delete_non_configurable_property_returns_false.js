// vybe-test: js/operators_deep/delete_non_configurable_property_returns_false
// origin: languages/js/tests/js/test_operators_deep.rs

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

const obj = Object.defineProperty({}, "x", { value: 1, configurable: false });
__check(__line(delete obj.x), "false");
__check(__line(obj.x), "1");
