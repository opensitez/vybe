// vybe-test: js/dynamic_programming/count_combinations
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

function combinations(n, k) {
    if (k === 0 || k === n) return 1;
    const dp = Array.from({length: n+1}, (_, i) => Array(i+1).fill(1));
    for (let i = 2; i <= n; i++) {
        for (let j = 1; j < i; j++) {
            dp[i][j] = dp[i-1][j-1] + dp[i-1][j];
        }
    }
    return dp[n][k];
}
console.log(combinations(4, 2));
console.log(combinations(5, 0));
console.log(combinations(6, 3));
