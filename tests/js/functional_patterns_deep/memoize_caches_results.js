// vybe-test: js/functional_patterns_deep/memoize_caches_results
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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
    return function(...args) {
        const key = JSON.stringify(args);
        if (cache.has(key)) return cache.get(key);
        const result = fn.apply(this, args);
        cache.set(key, result);
        return result;
    };
}
let calls = 0;
const expensiveFn = memoize(x => { calls++; return x * x; });
__check(__line(expensiveFn(5)), "25");
__check(__line(expensiveFn(5)), "25");
__check(__line(expensiveFn(6)), "36");
__check(__line(calls), "2"); // only 2: 5 and 6
