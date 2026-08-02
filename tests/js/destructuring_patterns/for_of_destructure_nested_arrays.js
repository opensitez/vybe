// vybe-test: js/destructuring_patterns/for_of_destructure_nested_arrays
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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
const sums = [];
for (const [a, b] of matrix) sums.push(a + b);
console.log(sums.join(","));
