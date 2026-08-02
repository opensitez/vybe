// vybe-test: js/array_higher_order/array_unique_via_set
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

const arr = [1, 2, 2, 3, 1, 4, 3];
const unique = [...new Set(arr)];
__check(__line(unique.join(",")), "1,2,3,4");
