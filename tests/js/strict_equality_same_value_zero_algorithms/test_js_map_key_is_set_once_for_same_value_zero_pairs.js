// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_map_key_is_set_once_for_same_value_zero_pairs
// origin: languages/js/tests/js/test_js_strict_equality_same_value_zero_algorithms.rs

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
map.set(+0, "plus");
map.set(-0, "minus");
map.set(NaN, "first");
map.set(NaN, "second");
__check(__line(`${map.size}:${map.get(0)}:${map.get(-0)}:${map.get(NaN)}`), "2:minus:minus:second");
