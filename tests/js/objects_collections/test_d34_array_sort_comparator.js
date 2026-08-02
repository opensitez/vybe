// vybe-test: js/objects_collections/test_d34_array_sort_comparator
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

let arr = [3, 1, 4, 1, 5, 9];
        arr.sort((a, b) => a - b);
        __check(__line(arr.join(",")), "1,1,3,4,5,9");
