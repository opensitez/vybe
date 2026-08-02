// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_array_index_as_string
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

const { ["0"]: first, ["1"]: second } = [100, 200];
__check(__line(`${first}:${second}`), "100:200");
