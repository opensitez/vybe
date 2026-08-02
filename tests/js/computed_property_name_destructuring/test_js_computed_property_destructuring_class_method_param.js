// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_class_method_param
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

class Handler {
    process(key, { [key]: val }) {
        return val * 10;
    }
}
__check(__line(new Handler().process("amount", { amount: 5 })), "50");
