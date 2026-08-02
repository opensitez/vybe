// vybe-test: js/dynamic_programming/memoized_fibonacci
// origin: languages/js/tests/js/test_dynamic_programming.rs

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

const memo = {};
function fib(n) {
    if (n in memo) return memo[n];
    if (n <= 1) return n;
    return memo[n] = fib(n-1) + fib(n-2);
}
__check(__line(fib(30)), "832040");
__check(__line(fib(35)), "9227465");
