// vybe-test: js/es2023_2025_features/array_find_last
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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
__check(__line(arr.findLast(x => x % 2 === 0)), "4");
__check(__line(arr.findLastIndex(x => x % 2 === 0)), "3");
