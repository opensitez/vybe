// vybe-test: js/map_set_deep_patterns/map_as_multimap
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

class MultiMap {
    #map = new Map();
    add(key, value) {
        if (!this.#map.has(key)) this.#map.set(key, []);
        this.#map.get(key).push(value);
        return this;
    }
    get(key) { return this.#map.get(key) ?? []; }
    has(key) { return this.#map.has(key); }
}
const mm = new MultiMap();
mm.add("a", 1).add("b", 2).add("a", 3).add("a", 4);
__check(__line(mm.get("a").join(",")), "1,3,4");
__check(__line(mm.get("b").join(",")), "2");
__check(__line(mm.has("c")), "false");
