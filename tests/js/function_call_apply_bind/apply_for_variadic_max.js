// vybe-test: js/function_call_apply_bind/apply_for_variadic_max
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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

const nums = [3, 1, 4, 1, 5, 9, 2, 6];
__check(__line(Math.max.apply(null, nums)), "9");
// Equivalent with spread:
__check(__line(Math.max(...nums)), "9");
