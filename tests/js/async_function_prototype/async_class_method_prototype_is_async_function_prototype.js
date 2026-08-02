// vybe-test: js/async_function_prototype/async_class_method_prototype_is_async_function_prototype
// origin: languages/js/tests/js/test_async_function_prototype.rs

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

class S { async load() {} } __check(__line(Object.getPrototypeOf(S.prototype.load) === AsyncFunction.prototype), "true");
