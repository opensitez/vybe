// vybe-test: js/class_decorators/method_decorator_wraps_method
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

function logged(fn, name) {
    return function(...args) {
        __check(__line("call:" + name), "call:add");
        return fn.apply(this, args);
    };
}
class Calculator {
    add(a, b) { return a + b; }
}
Calculator.prototype.add = logged(Calculator.prototype.add, "add");
const c = new Calculator();
__check(__line(c.add(2, 3)), "5");
