// vybe-test: js/function_prototype_metadata/arrow_length_counts_params_without_default
// origin: languages/js/tests/js/test_function_prototype_metadata.rs

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

const f = (a, b, c) => {}; __check(__line(f.length), "3");
