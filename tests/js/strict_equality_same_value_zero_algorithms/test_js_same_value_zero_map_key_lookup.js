// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_same_value_zero_map_key_lookup
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
map.set(+0, "zero");
map.set(NaN, "not_a_number");

__check(__line(`${map.get(-0)}:${map.get(NaN)}`), "zero:not_a_number");
