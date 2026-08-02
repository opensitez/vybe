// vybe-test: js/module_patterns/memoize_with_max_size
// origin: languages/js/tests/js/test_module_patterns.rs

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

function lruMemoize(fn, maxSize = 3) {
    const cache = new Map();
    return function(key) {
        if (cache.has(key)) {
            const val = cache.get(key);
            cache.delete(key);
            cache.set(key, val); // move to end (most recent)
            return val;
        }
        const result = fn(key);
        if (cache.size >= maxSize) {
            cache.delete(cache.keys().next().value); // remove oldest
        }
        cache.set(key, result);
        return result;
    };
}
let calls = 0;
const sq = lruMemoize(x => { calls++; return x * x; }, 2);
sq(2); sq(3); sq(2); sq(4); // sq(4) evicts sq(3)
sq(3); // must recompute (evicted)
__check(__line(calls), "5"); // 2+3+4+3 computed = 4 unique + 1 recompute = 5
