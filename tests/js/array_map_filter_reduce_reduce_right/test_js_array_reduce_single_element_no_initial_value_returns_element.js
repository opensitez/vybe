// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_reduce_single_element_no_initial_value_returns_element
// origin: languages/js/tests/js/test_js_array_map_filter_reduce_reduce_right.rs

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

const val = [42].reduce((acc, curr) => acc + curr);
__check(__line(val), "42");
