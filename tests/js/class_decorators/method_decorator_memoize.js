// vybe-test: js/class_decorators/method_decorator_memoize
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

function memoize(fn) {
    const cache = new Map();
    return function(x) {
        if (cache.has(x)) { __check(__line("cached"), "16"); return cache.get(x); }
        const result = fn.call(this, x);
        cache.set(x, result);
        return result;
    };
}
class Math2 {
    square(n) { return n * n; }
}
Math2.prototype.square = memoize(Math2.prototype.square);
const m = new Math2();
__check(__line(m.square(4)), "cached");
__check(__line(m.square(4)), "16");
