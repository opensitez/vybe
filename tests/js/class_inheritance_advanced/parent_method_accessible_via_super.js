// vybe-test: js/class_inheritance_advanced/parent_method_accessible_via_super
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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
    log(msg) { return "[LOG] " + msg; }
}
class PrefixLogger extends Logger {
    constructor(prefix) {
        super();
        this.prefix = prefix;
    }
    log(msg) {
        return super.log(this.prefix + ": " + msg);
    }
}
const logger = new PrefixLogger("App");
__check(__line(logger.log("started")), "[LOG] App: started");
