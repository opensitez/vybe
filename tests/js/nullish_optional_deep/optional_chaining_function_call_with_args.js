// vybe-test: js/nullish_optional_deep/optional_chaining_function_call_with_args
// origin: languages/js/tests/js/test_nullish_optional_deep.rs

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

const obj = {
    greet: (name) => "Hello " + name
};
__check(__line(obj.greet?.("World")), "Hello World");
__check(__line(obj.missing?.("World")), "undefined");
