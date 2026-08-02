// vybe-test: js/dataview_typed_array_deep/typed_array_from_array
// origin: languages/js/tests/js/test_dataview_typed_array_deep.rs

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

const arr = Int32Array.from([1, 2, 3, 4]);
__check(__line(arr[0]), "1");
__check(__line(arr.length), "4");
__check(__line(arr instanceof Int32Array), "true");
