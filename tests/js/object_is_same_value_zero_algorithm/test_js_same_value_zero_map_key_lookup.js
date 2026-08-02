// vybe-test: js/object_is_same_value_zero_algorithm/test_js_same_value_zero_map_key_lookup
// origin: languages/js/tests/js/test_js_object_is_same_value_zero_algorithm.rs

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
map.set(NaN, "NaN_Value");
map.set(-0, "Zero_Value");

__check(__line(map.get(NaN)), "NaN_Value");
__check(__line(map.get(+0)), "Zero_Value");
