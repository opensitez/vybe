// vybe-test: js/math_algorithms/clamp_lerp_normalize
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

const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
const lerp = (a, b, t) => a + (b - a) * t;
const normalize = (v, min, max) => (v - min) / (max - min);
__check(__line(clamp(150, 0, 100)), "100");
__check(__line(lerp(0, 100, 0.5)), "50");
__check(__line(normalize(75, 0, 100)), "0.75");
__check(__line(clamp(-5, 0, 100)), "0");
