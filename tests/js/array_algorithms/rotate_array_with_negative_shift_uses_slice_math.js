// vybe-test: js/array_algorithms/rotate_array_with_negative_shift_uses_slice_math
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function rotateLeft(arr, k) {
    const n = arr.length;
    k = k % n;
    return [...arr.slice(k), ...arr.slice(0, k)];
}
__check(__line(rotateLeft([1, 2, 3], -1).join(",")), "3,1,2");
