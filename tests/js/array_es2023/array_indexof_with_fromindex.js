// vybe-test: js/array_es2023/array_indexof_with_fromindex
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

const arr = [1, 2, 3, 2, 1];
__check(__line(arr.indexOf(2)), "1");
__check(__line(arr.indexOf(2, 2)), "3");
__check(__line(arr.indexOf(2, 4)), "-1");
