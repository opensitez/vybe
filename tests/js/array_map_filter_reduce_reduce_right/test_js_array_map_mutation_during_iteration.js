// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_map_mutation_during_iteration
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

const arr = [1, 2, 3];
const res = arr.map((x, idx, a) => {
    if (idx === 0) a[2] = 99; // Mutates original array element before visited
    return x;
});
__check(__line(res.join(",")), "1,2,99");
