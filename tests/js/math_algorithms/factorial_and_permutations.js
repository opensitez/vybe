// vybe-test: js/math_algorithms/factorial_and_permutations
// origin: languages/js/tests/js/test_math_algorithms.rs

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

function factorial(n) { return n <= 1 ? 1 : n * factorial(n - 1); }
function permutations(n, r) { return factorial(n) / factorial(n - r); }
function combinations(n, r) { return permutations(n, r) / factorial(r); }
__check(__line(factorial(5)), "120");
__check(__line(permutations(5, 2)), "20");
__check(__line(combinations(5, 2)), "10");
