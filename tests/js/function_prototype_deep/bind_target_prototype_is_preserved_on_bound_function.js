// vybe-test: js/function_prototype_deep/bind_target_prototype_is_preserved_on_bound_function
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

function decl() {} const b = decl.bind(null); __check(__line(Object.getPrototypeOf(b) === Function.prototype), "true");
