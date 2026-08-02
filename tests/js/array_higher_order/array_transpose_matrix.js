// vybe-test: js/array_higher_order/array_transpose_matrix
// origin: languages/js/tests/js/test_array_higher_order.rs

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

const matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
const transposed = matrix[0].map((_, col) => matrix.map(row => row[col]));
__check(__line(transposed[0].join(",")), "1,4,7");
__check(__line(transposed[1].join(",")), "2,5,8");
