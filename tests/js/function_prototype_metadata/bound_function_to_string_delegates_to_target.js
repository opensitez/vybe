// vybe-test: js/function_prototype_metadata/bound_function_to_string_delegates_to_target
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

function orig() { return 0; } const b = orig.bind(null); __check(__line(Function.prototype.toString.call(b) === Function.prototype.toString.call(orig)), "false");
