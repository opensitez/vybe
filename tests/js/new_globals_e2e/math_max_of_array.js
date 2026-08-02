// vybe-test: js/new_globals_e2e/math_max_of_array
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

__check(__line(Math.maxOf([3, 1, 4, 1, 5, 9, 2, 6])), "9");
