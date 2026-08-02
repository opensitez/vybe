// vybe-test: js/class_decorators/class_decorator_replaces_class
// origin: languages/js/tests/js/test_class_decorators.rs

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

function withVersion(target) { target.version = "1.0"; return target; }
class App {}
withVersion(App);
__check(__line(App.version), "1.0");
