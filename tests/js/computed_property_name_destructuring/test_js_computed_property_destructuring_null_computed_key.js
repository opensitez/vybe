// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_null_computed_key
// origin: languages/js/tests/js/test_js_computed_property_name_destructuring.rs

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

const nullKey = null;
const { [nullKey]: val } = { null: "NullKeyVal" };
__check(__line(val), "NullKeyVal");
