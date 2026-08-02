// vybe-test: js/array_es2023/nested_array_at_method
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

const matrix = [[1, 2], [3, 4], [5, 6]];
__check(__line(matrix.at(0).at(-1)), "2");
__check(__line(matrix.at(-1).at(0)), "5");
