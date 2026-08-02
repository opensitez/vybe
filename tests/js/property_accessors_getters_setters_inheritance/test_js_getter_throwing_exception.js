// vybe-test: js/property_accessors_getters_setters_inheritance/test_js_getter_throwing_exception
// origin: languages/js/tests/js/test_js_property_accessors_getters_setters_inheritance.rs

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

const obj = {
    get fail() { throw new Error("GetterFailed"); }
};
try {
    const val = obj.fail;
} catch (e) {
    __check(__line(e.message), "GetterFailed");
}
