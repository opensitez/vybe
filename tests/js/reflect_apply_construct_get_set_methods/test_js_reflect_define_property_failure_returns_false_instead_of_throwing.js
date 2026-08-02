// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_define_property_failure_returns_false_instead_of_throwing
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set_methods.rs

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
Object.defineProperty(obj, "locked", { value: 1, configurable: false });
const success = Reflect.defineProperty(obj, "locked", { configurable: true });
__check(__line(success), "false"); // Returns false cleanly without throwing exception!
