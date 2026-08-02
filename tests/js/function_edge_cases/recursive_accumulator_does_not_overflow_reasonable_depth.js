// vybe-test: js/function_edge_cases/recursive_accumulator_does_not_overflow_reasonable_depth
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

function sum(n, acc = 0) {
    if (n <= 0) return acc;
    return sum(n - 1, acc + n);
}
__check(__line(sum(100)), "5050");
