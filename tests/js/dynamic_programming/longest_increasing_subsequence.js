// vybe-test: js/dynamic_programming/longest_increasing_subsequence
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

function lis(arr) {
    const dp = Array(arr.length).fill(1);
    for (let i = 1; i < arr.length; i++) {
        for (let j = 0; j < i; j++) {
            if (arr[j] < arr[i]) dp[i] = Math.max(dp[i], dp[j] + 1);
        }
    }
    return Math.max(...dp);
}
console.log(lis([10, 9, 2, 5, 3, 7, 101, 18]));
console.log(lis([0, 1, 0, 3, 2, 3]));
