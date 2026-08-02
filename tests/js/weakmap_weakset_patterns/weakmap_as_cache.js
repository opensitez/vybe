// vybe-test: js/weakmap_weakset_patterns/weakmap_as_cache
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const cache = new WeakMap();
function process(obj) {
    if (cache.has(obj)) return cache.get(obj);
    const result = Object.keys(obj).length;
    cache.set(obj, result);
    return result;
}
const o = { a: 1, b: 2, c: 3 };
__check(__line(process(o)), "3");
__check(__line(process(o)), "3"); // from cache
