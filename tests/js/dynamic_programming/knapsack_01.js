// vybe-test: js/dynamic_programming/knapsack_01
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

function knapsack(weights, values, capacity) {
    const n = weights.length;
    const dp = Array.from({length: n+1}, () => Array(capacity+1).fill(0));
    for (let i = 1; i <= n; i++) {
        for (let w = 0; w <= capacity; w++) {
            dp[i][w] = dp[i-1][w];
            if (weights[i-1] <= w) {
                dp[i][w] = Math.max(dp[i][w], dp[i-1][w-weights[i-1]] + values[i-1]);
            }
        }
    }
    return dp[n][capacity];
}
console.log(knapsack([2, 3, 4, 5], [3, 4, 5, 6], 5));
