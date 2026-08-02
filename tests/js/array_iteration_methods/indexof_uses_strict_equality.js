// vybe-test: js/array_iteration_methods/indexof_uses_strict_equality
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

const arr = [1, 2, NaN, 3];
__check(__line(arr.indexOf(NaN)), "-1"); // -1 — NaN !== NaN
__check(__line(arr.indexOf(2)), "1");
