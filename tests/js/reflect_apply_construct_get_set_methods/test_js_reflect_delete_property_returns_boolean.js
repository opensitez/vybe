// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_delete_property_returns_boolean
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

const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, configurable: false });
__check(__line(Reflect.deleteProperty(obj, "a")), "true");
__check(__line(Reflect.deleteProperty(obj, "b")), "false");
__check(__line("a" in obj), "false");
