// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_map_length_fixed_at_start
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

const arr = [1, 2];
const res = arr.map((x, idx, a) => {
    if (idx === 0) a.push(3); // Pushed element should NOT be visited by map!
    return x * 10;
});
__check(__line(res.join(",") + "|arrLength=" + arr.length), "10,20|arrLength=3");
