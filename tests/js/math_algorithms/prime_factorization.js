// vybe-test: js/math_algorithms/prime_factorization
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

function primeFactors(n) {
    const factors = [];
    for (let d = 2; d * d <= n; d++) {
        while (n % d === 0) { factors.push(d); n /= d; }
    }
    if (n > 1) factors.push(n);
    return factors;
}
console.log(primeFactors(12).join(","));
console.log(primeFactors(100).join(","));
console.log(primeFactors(17).join(","));
