// vybe-test: js/class_constructor_errors/derived_instanceof_base_after_successful_construct
// origin: languages/js/tests/js/test_class_constructor_errors.rs

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

class B{} class D extends B{} const d=new D();__check(__line(d instanceof B), "true");__check(__line(d instanceof D), "true");
