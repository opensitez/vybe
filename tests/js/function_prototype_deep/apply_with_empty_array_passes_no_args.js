// vybe-test: js/function_prototype_deep/apply_with_empty_array_passes_no_args
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

function len() { return arguments.length; } __check(__line(apply_len()), "0"); function apply_len() { return len.apply(null, []); }
