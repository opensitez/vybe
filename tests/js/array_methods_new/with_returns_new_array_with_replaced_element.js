// vybe-test: js/array_methods_new/with_returns_new_array_with_replaced_element
// origin: languages/js/tests/js/test_array_methods_new.rs

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

const arr = [1, 2, 3, 4];
const result = arr.with(2, 99);
__check(__line(result.join(",")), "1,2,99,4");
__check(__line(arr[2]), "3"); // original unchanged
