// vybe-test: js/map_set_deep/weakmap_stores_per_object_data
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

const meta = new WeakMap();
const obj1 = {};
const obj2 = {};
meta.set(obj1, { created: 2024 });
meta.set(obj2, { created: 2025 });
__check(__line(meta.get(obj1).created), "2024");
__check(__line(meta.get(obj2).created), "2025");
__check(__line(meta.has({})), "false"); // different object
