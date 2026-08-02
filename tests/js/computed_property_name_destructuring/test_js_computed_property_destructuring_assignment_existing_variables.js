// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_assignment_existing_variables
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

let val;
const key = "x";
({ [key]: val } = { x: 10 });
__check(__line(val), "10");
