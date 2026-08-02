// vybe-test: js/ecma_operators/delete_non_configurable_property_returns_false
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

const obj = {};
Object.defineProperty(obj, "x", {
    value: 42,
    writable: false,
    configurable: false,
    enumerable: true,
});

__check(__line(delete obj.x), "false");
__check(__line(obj.x), "42");
