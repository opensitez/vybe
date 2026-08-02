// vybe-test: js/weakref_finalization/weakref_cache_pattern
// origin: languages/js/tests/js/test_weakref_finalization.rs

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

// Simulate a cache that holds weak references to objects
class WeakCache {
    #map = new Map();

    set(key, value) {
        this.#map.set(key, new WeakRef(value));
    }

    get(key) {
        return this.#map.get(key)?.deref();
    }
}

const cache = new WeakCache();
const obj = { data: "important" };
cache.set("key", obj);
__check(__line(cache.get("key")?.data), "important");
