// vybe-test: js/functional_fp_patterns/memoize_recursive
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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
let calls = 0;
const square = memoize(n => { calls++; return n * n; });
__check(__line(square(5)), "25");
__check(__line(square(5)), "25");
__check(__line(square(3)), "9");
const c = calls;
square(5);
square(3);
__check(__line(calls === c), "true");
