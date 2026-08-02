// vybe-test: js/oop_patterns_advanced/flyweight_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class TreeType {
    constructor(name, color) { this.name=name; this.color=color; }
}
const _cache = new Map();
class TreeFactory {
    static get(name, color) {
        const key = name+color;
        if (!_cache.has(key)) _cache.set(key, new TreeType(name, color));
        return _cache.get(key);
    }
    static size() { return _cache.size; }
}
const t1 = TreeFactory.get("Oak", "green");
const t2 = TreeFactory.get("Oak", "green");
const t3 = TreeFactory.get("Pine", "dark");
__check(__line(t1 === t2), "true");
__check(__line(t1 === t3), "false");
__check(__line(TreeFactory.size()), "2");
