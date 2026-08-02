// vybe-test: js/ecma_arrays/map_filter_chain
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

const result = [1, 2, 3, 4, 5]
    .filter(x => x % 2 !== 0)
    .map(x => x * x);
__check(__line(result.join(",")), "1,9,25");
