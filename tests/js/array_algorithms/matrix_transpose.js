// vybe-test: js/array_algorithms/matrix_transpose
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function transpose(matrix) {
    return matrix[0].map((_, i) => matrix.map(row => row[i]));
}
const m = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
const t = transpose(m);
__check(__line(t[0].join(",")), "1,4,7");
__check(__line(t[1].join(",")), "2,5,8");
__check(__line(t[2].join(",")), "3,6,9");
