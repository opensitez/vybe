// vybe-test: js/array_algorithms/sliding_window_max
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function maxSubarraySum(arr, k) {
    let sum = arr.slice(0, k).reduce((a, b) => a + b, 0);
    let max = sum;
    for (let i = k; i < arr.length; i++) {
        sum += arr[i] - arr[i - k];
        max = Math.max(max, sum);
    }
    return max;
}
console.log(maxSubarraySum([1, 3, -1, -3, 5, 3, 6, 7], 3));
