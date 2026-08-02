// vybe-test: js/object_prototype_patterns/object_create_null_safe_map
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

const map = {};
map.key = "value";
map.other = "data";
__check(__line(map.key), "value");
__check(__line(Object.keys(map).length), "2");
// User-added keys are own; inherited ones like toString are not own
__check(__line(Object.prototype.hasOwnProperty.call(map, "toString")), "false");
