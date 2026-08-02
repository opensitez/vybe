// vybe-test: js/array_es2023/array_tosorted_map_filter_chain
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

const nums = [5, 2, 8, 1, 9, 3];
const result = nums
    .toSorted((a, b) => a - b)
    .filter(x => x > 3)
    .map(x => x * 10);
__check(__line(result.join(",")), "50,80,90");
