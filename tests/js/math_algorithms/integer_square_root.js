// vybe-test: js/math_algorithms/integer_square_root
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

function isqrt(n) {
    if (n < 0) return NaN;
    let x = Math.floor(Math.sqrt(n));
    while (x * x > n) x--;
    while ((x+1)*(x+1) <= n) x++;
    return x;
}
console.log(isqrt(16));
console.log(isqrt(17));
console.log(isqrt(100));
console.log(isqrt(0));
