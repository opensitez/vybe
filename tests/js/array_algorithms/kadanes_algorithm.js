// vybe-test: js/array_algorithms/kadanes_algorithm
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

function maxSubarray(arr) {
    let maxSum = arr[0], current = arr[0];
    for (let i = 1; i < arr.length; i++) {
        current = Math.max(arr[i], current + arr[i]);
        maxSum = Math.max(maxSum, current);
    }
    return maxSum;
}
console.log(maxSubarray([-2, 1, -3, 4, -1, 2, 1, -5, 4]));
console.log(maxSubarray([-1, -2, -3]));
