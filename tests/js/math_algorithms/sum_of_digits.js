// vybe-test: js/math_algorithms/sum_of_digits
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

function digitSum(n) {
    return Math.abs(n).toString().split("").reduce((s, d) => s + Number(d), 0);
}
function digitalRoot(n) {
    while (n > 9) n = digitSum(n);
    return n;
}
console.log(digitSum(12345));
console.log(digitalRoot(9875));
console.log(digitalRoot(0));
