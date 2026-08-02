// vybe-test: js/getter_setter_deep/accessor_descriptor_does_not_have_value
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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

const obj = { get x() { return 1; } };
const desc = Object.getOwnPropertyDescriptor(obj, "x");
__check(__line("value" in desc), "false");
__check(__line(typeof desc.get), "function");
__check(__line(typeof desc.set), "undefined");
__check(__line(desc.enumerable), "true");
