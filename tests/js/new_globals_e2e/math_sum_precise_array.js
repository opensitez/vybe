// vybe-test: js/new_globals_e2e/math_sum_precise_array
// origin: languages/js/tests/js/test_new_globals_e2e.rs

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

__check(__line(Math.sumPrecise([1, 2, 3, 4, 5])), "15");
