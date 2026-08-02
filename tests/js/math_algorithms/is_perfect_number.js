// vybe-test: js/math_algorithms/is_perfect_number
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

function isPerfect(n) {
    if (n <= 1) return false;
    let sum = 1;
    for (let i = 2; i * i <= n; i++) {
        if (n % i === 0) { sum += i; if (i !== n/i) sum += n/i; }
    }
    return sum === n;
}
console.log(isPerfect(6));
console.log(isPerfect(28));
console.log(isPerfect(12));
