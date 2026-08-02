// vybe-test: js/function_prototype_deep/apply_with_array_like_object_with_length
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

function first(a) { return a; } const like = { 0: "z", length: 1 }; __check(__line(first.apply(null, like)), "z");
