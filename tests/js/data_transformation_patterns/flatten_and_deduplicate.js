// vybe-test: js/data_transformation_patterns/flatten_and_deduplicate
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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

const nested = [[1, 2, 3], [2, 3, 4], [4, 5]];
const unique = [...new Set(nested.flat())].sort((a, b) => a - b);
__check(__line(unique.join(",")), "1,2,3,4,5");
