// vybe-test: js/function_prototype_deep/apply_passes_array_elements_not_array_itself
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

function isArray(a) { return Array.isArray(a); } __check(__line(isArray.apply(null, [[1]])), "true");
