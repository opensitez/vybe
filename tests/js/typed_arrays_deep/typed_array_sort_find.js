// vybe-test: js/typed_arrays_deep/typed_array_sort_find
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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

const arr = new Int32Array([5, 3, 1, 4, 2]);
arr.sort();
__check(__line(Array.from(arr).join(",")), "1,2,3,4,5");
__check(__line(arr.find(x => x > 3)), "4");
__check(__line(arr.findIndex(x => x > 3)), "3");
