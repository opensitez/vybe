// vybe-test: js/inheritance/test_15_override_calls_super_then_extends
// origin: languages/js/tests/js/js_inheritance_test.rs

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
            format(msg) { return "[LOG] " + msg; }
        }
        class TimedLogger extends Logger {
            constructor() { super(); }
            format(msg) { return super.format(msg) + " @now"; }
        }
        let tl = new TimedLogger();
        __check(__line(tl.format("hi")), "[LOG] hi @now");
