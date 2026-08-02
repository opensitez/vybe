// vybe-test: js/map_set_deep_patterns/map_object_keys
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

const map = new Map();
const key1 = { id: 1 };
const key2 = { id: 2 };
map.set(key1, "value1");
map.set(key2, "value2");
__check(__line(map.get(key1)), "value1");
__check(__line(map.get(key2)), "value2");
__check(__line(map.size), "2");
__check(__line(map.has({ id: 1 })), "false");
