// vybe-test: js/class_super_method_property_access/test_js_class_super_static_method_call
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

class Logger {
    static log(msg) { return `[LOG] ${msg}`; }
}
class CustomLogger extends Logger {
    static log(msg) {
        return super.log(msg).toUpperCase();
    }
}
__check(__line(CustomLogger.log("Hello")), "[LOG] HELLO");
