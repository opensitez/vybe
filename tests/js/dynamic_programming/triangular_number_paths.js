// vybe-test: js/dynamic_programming/triangular_number_paths
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

function minTrianglePath(triangle) {
    const dp = [...triangle[triangle.length - 1]];
    for (let i = triangle.length - 2; i >= 0; i--) {
        for (let j = 0; j <= i; j++) {
            dp[j] = triangle[i][j] + Math.min(dp[j], dp[j+1]);
        }
    }
    return dp[0];
}
const t = [[2], [3, 4], [6, 5, 7], [4, 1, 8, 3]];
console.log(minTrianglePath(t));
