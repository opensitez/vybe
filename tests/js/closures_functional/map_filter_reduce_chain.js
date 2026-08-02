// vybe-test: js/closures_functional/map_filter_reduce_chain
// origin: languages/js/tests/js/test_closures_functional.rs

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

let data = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
let result = data
    .filter(x => x % 2 === 0)
    .map(x => x * x)
    .reduce((acc, x) => acc + x, 0);
__check(__line(result), "220");
