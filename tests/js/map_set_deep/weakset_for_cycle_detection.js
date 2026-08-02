// vybe-test: js/map_set_deep/weakset_for_cycle_detection
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

function hasCycle(obj, seen = new WeakSet()) {
    if (typeof obj !== "object" || obj === null) return false;
    if (seen.has(obj)) return true;
    seen.add(obj);
    return Object.values(obj).some(v => hasCycle(v, seen));
}

const normal = { a: { b: { c: 1 } } };
__check(__line(hasCycle(normal)), "false");

const cyclic = {};
cyclic.self = cyclic;
__check(__line(hasCycle(cyclic)), "true");
