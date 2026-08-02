// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_undefined_computed_key_throws
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

let missingVar;
try {
    const { [missingVar]: val } = { undefined: "UndefinedKeyVal" };
    console.log(val); // In JS, undefined is converted to "undefined" key string!
} catch (e) {
    console.log("Error");
}
