// vybe-test: js/string_array_advanced/array_splice_remove_and_insert
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let arr = [1, 2, 3, 4, 5];
let removed = arr.splice(1, 2, 10, 20, 30);
__check(__line(removed.join(",")), "2,3");
__check(__line(arr.join(",")), "1,10,20,30,4,5");
