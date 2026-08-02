// vybe-test: js/object_methods_deep/object_from_entries_from_map
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const map = new Map([["key1", "val1"], ["key2", "val2"]]);
const obj = Object.fromEntries(map);
__check(__line(obj.key1), "val1");
__check(__line(obj.key2), "val2");
