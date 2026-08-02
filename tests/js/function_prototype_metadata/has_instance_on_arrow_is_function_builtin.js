// vybe-test: js/function_prototype_metadata/has_instance_on_arrow_is_function_builtin
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

const arrow = () => {}; __check(__line(Function[Symbol.hasInstance](arrow)), "true");
