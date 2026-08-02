// vybe-test: js/map_set_deep/map_uses_reference_equality_for_object_keys
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

const key1 = { id: 1 };
const key2 = { id: 1 }; // different object, same content
const m = new Map();
m.set(key1, "value1");
__check(__line(m.has(key1)), "true");
__check(__line(m.has(key2)), "false"); // different reference
__check(__line(m.size), "1");
