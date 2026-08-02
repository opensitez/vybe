// vybe-test: js/array_sort_advanced/sort_sparse_array_preserves_length_and_moves_hole_to_end
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const arr = ["b", "a", , undefined];
arr.sort();
__check(__line(arr.length), "4");
__check(__line(0 in arr), "true");
__check(__line(2 in arr), "true");
__check(__line(3 in arr), "true");
__check(__line(arr.join(",")), "a,b,,");
