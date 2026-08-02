// vybe-test: js/array_higher_order/array_sliding_window
// origin: languages/js/tests/js/test_array_higher_order.rs

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

function windows(arr, size) {
    return arr.slice(0, arr.length - size + 1).map((_, i) => arr.slice(i, i + size));
}
const result = windows([1, 2, 3, 4, 5], 3);
__check(__line(result.length), "3");
__check(__line(result[0].join(",")), "1,2,3");
__check(__line(result[2].join(",")), "3,4,5");
