// vybe-test: js/function_prototype_deep/apply_with_sparse_array_preserves_holes_as_undefined
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function pair(a, b) { return a === undefined && b === undefined; } const sparse = [1, , 3]; __check(__line(pair.apply(null, sparse)), "false");
