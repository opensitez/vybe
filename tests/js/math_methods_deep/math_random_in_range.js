// vybe-test: js/math_methods_deep/math_random_in_range
// origin: languages/js/tests/js/test_math_methods_deep.rs

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

// Can't assert exact value, but range and type
const r = Math.random();
console.log(typeof r);
console.log(r >= 0 && r < 1);
