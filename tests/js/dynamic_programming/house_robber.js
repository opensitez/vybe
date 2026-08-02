// vybe-test: js/dynamic_programming/house_robber
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

function rob(nums) {
    let prev2 = 0, prev1 = 0;
    for (const n of nums) {
        const curr = Math.max(prev1, prev2 + n);
        prev2 = prev1;
        prev1 = curr;
    }
    return prev1;
}
console.log(rob([1, 2, 3, 1]));
console.log(rob([2, 7, 9, 3, 1]));
