// vybe-test: js/closures_functional/memoize_fibonacci
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
let fib = memoize(function(n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
});
__check(__line(fib(10)), "55");
__check(__line(fib(20)), "6765");
