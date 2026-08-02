// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_same_value_zero_used_in_map_and_set
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

const set = new Set();
set.add(+0);
set.add(-0);
set.add(NaN);
set.add(NaN);

__check(__line(`${set.size}:${set.has(-0)}:${set.has(NaN)}`), "2:true:true");
