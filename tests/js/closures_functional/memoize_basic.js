// vybe-test: js/closures_functional/memoize_basic
// origin: languages/js/tests/js/test_closures_functional.rs

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
    let cache = {};
    return function(n) {
        if (cache[n] !== undefined) return cache[n];
        cache[n] = fn(n);
        return cache[n];
    };
}
let callCount = 0;
let square = memoize(n => { callCount++; return n * n; });
__check(__line(square(4)), "16");
__check(__line(square(4)), "16");
__check(__line(square(5)), "25");
__check(__line(callCount), "2");
