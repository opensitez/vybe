// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_map_index_and_array_arguments
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

const letters = ["a", "b"];
const res = letters.map((val, idx, arr) => `${val}:${idx}:${arr.length}`);
__check(__line(res.join("|")), "a:0:2|b:1:2");
