// vybe-test: js/map_set_deep/map_as_lru_like_cache
// origin: languages/js/tests/js/test_map_set_deep.rs

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
let computeCount = 0;

function compute(key) {
    if (cache.has(key)) return cache.get(key);
    computeCount++;
    const result = key * key;
    cache.set(key, result);
    return result;
}

compute(5); compute(5); compute(6); compute(5);
__check(__line(computeCount), "2");  // 2 unique computations
__check(__line(cache.get(5)), "25");
