// vybe-test: js/math_algorithms/vector_normalization_hypot
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

const norm = (x, y) => { const len = Math.hypot(x, y); return [x / len, y / len]; };
__check(__line(norm(3, 4).join(",")), "0.6,0.8");
