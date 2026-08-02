// vybe-test: js/spread_rest_advanced/object_spread_nested_shallow
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

const a = { nested: { val: 1 } };
const b = { ...a };
b.nested.val = 99;
__check(__line(a.nested.val), "99");
