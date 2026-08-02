// vybe-test: js/function_prototype_deep/function_declaration_has_prototype_object
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

function ctor() {} __check(__line(typeof ctor.prototype === "object"), "true");
