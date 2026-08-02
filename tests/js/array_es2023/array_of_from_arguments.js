// vybe-test: js/array_es2023/array_of_from_arguments
// origin: languages/js/tests/js/test_array_es2023.rs

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

const arr = Array.of(7, 8, 9);
__check(__line(arr.length), "3");
__check(__line(arr.join(",")), "7,8,9");
