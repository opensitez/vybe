// vybe-test: js/array_iteration_methods/reduce_no_initial_uses_first_element
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

const max = [3, 1, 4, 1, 5, 9].reduce((a, b) => a > b ? a : b);
__check(__line(max), "9");
