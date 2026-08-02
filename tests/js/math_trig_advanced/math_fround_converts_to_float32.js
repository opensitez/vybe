// vybe-test: js/math_trig_advanced/math_fround_converts_to_float32
// origin: languages/js/tests/js/test_math_trig_advanced.rs

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

// 1.337 can't be exactly represented in float32
const x = Math.fround(1.337);
console.log(x !== 1.337);       // precision differs
console.log(Math.fround(0));
console.log(Math.fround(1));
