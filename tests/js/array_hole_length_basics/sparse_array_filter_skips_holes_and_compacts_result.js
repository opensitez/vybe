// vybe-test: js/array_hole_length_basics/sparse_array_filter_skips_holes_and_compacts_result
// origin: languages/js/tests/js/test_array_hole_length_basics.rs

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

const arr = [, 1, , 2];
const filtered = arr.filter(() => true);
__check(__line(filtered.length), "2");
__check(__line(filtered.join(",")), "1,2");
