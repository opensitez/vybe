// vybe-test: js/array_iteration_methods/findindex_returns_first_matching_index
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

__check(__line([1, 2, 3, 4].findIndex(n => n > 2)), "2");
__check(__line([1, 2, 3].findIndex(n => n > 10)), "-1");
