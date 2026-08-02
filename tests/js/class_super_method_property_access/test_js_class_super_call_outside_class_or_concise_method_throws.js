// vybe-test: js/class_super_method_property_access/test_js_class_super_call_outside_class_or_concise_method_throws
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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

try {
    eval("function standalone() { super.foo(); } standalone();");
} catch (e) {
    __check(__line("Super Outside Method Error"), "Super Outside Method Error");
}
