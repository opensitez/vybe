// vybe-test: js/array_iteration_methods/entries_yields_index_value_pairs
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

const pairs = [...["a", "b", "c"].entries()];
__check(__line(pairs.map(([i, v]) => i + ":" + v).join(",")), "0:a,1:b,2:c");
