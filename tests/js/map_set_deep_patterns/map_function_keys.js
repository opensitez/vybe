// vybe-test: js/map_set_deep_patterns/map_function_keys
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

const cache = new Map();
function memoize(fn) {
    return function(...args) {
        const key = JSON.stringify(args);
        if (!cache.has(key)) cache.set(key, fn(...args));
        return cache.get(key);
    };
}
const add = memoize((a, b) => a + b);
__check(__line(add(1, 2)), "3");
__check(__line(add(1, 2)), "3");
__check(__line(cache.size), "1");
