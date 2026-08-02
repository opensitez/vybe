// vybe-test: js/closure_scope_deep_patterns/closure_memoization_cache
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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
    return (...args) => {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
}
let callCount = 0;
const expensive = memoize((a, b) => { callCount++; return a + b; });
__check(__line(expensive(1, 2)), "3");
__check(__line(expensive(1, 2)), "3");
__check(__line(expensive(3, 4)), "7");
__check(__line(callCount), "2");
