// vybe-test: js/class_decorators/static_method_decorator
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

function once(fn) {
    let called = false, result;
    return function(...args) {
        if (!called) { called = true; result = fn.apply(this, args); }
        return result;
    };
}
class Factory {
    static count = 0;
    static create() { return ++Factory.count; }
}
Factory.create = once(Factory.create.bind(Factory));
__check(__line(Factory.create()), "1");
__check(__line(Factory.create()), "1");
__check(__line(Factory.count), "1");
