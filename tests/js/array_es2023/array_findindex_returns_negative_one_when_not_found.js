// vybe-test: js/array_es2023/array_findindex_returns_negative_one_when_not_found
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

const arr = [1, 2, 3, 4, 5];
__check(__line(arr.findIndex(x => x > 100)), "-1");
__check(__line(arr.findIndex(x => x === 0)), "-1");
