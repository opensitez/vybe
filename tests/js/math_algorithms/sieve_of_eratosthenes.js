// vybe-test: js/math_algorithms/sieve_of_eratosthenes
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

function sieve(n) {
    const primes = Array(n + 1).fill(true);
    primes[0] = primes[1] = false;
    for (let i = 2; i * i <= n; i++) {
        if (primes[i]) for (let j = i*i; j <= n; j+=i) primes[j] = false;
    }
    return primes.map((p,i)=>p?i:-1).filter(n=>n>0);
}
const p = sieve(30);
console.log(p.join(","));
