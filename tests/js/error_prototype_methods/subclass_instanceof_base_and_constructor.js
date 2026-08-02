// vybe-test: js/error_prototype_methods/subclass_instanceof_base_and_constructor
// origin: languages/js/tests/js/test_error_prototype_methods.rs

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

class E extends RangeError {} const e=new E();__check(__line(e instanceof RangeError), "true");__check(__line(e instanceof Error), "true");
