// vybe-test: js/mixin_abstract_patterns/test_mixin_overriding_static_method_with_super
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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

const StaticLogger = Base => class extends Base {
    static log(msg) { return "[STATIC] " + msg; }
};
class BaseLogger {}
class AppLogger extends StaticLogger(BaseLogger) {}
__check(__line(AppLogger.log("Ready")), "[STATIC] Ready");
