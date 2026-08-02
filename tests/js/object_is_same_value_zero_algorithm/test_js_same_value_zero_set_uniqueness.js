// vybe-test: js/object_is_same_value_zero_algorithm/test_js_same_value_zero_set_uniqueness
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

const set = new Set();
set.add(NaN);
set.add(NaN);
set.add(+0);
set.add(-0);

__check(__line(set.size), "2");
