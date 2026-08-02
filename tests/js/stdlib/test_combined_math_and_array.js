// vybe-test: js/stdlib/test_combined_math_and_array
// origin: languages/js/tests/js/js_stdlib_test.rs

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

let nums = [3.7, 1.2, 5.9, 2.1];
        let sum = 0;
        for (let i = 0; i < nums.length; i++) {
            sum = sum + Math.floor(nums[i]);
        }
        console.log(sum);
