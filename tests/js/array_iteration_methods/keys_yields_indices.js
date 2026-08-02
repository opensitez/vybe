// vybe-test: js/array_iteration_methods/keys_yields_indices
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

const keys = [...["a", "b", "c"].keys()];
__check(__line(keys.join(",")), "0,1,2");
