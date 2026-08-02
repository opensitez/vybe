// vybe-test: js/class_constructor_errors/constructor_with_rest_param_receives_args
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

class C{constructor(first,...rest){__check(__line(first), "1");__check(__line(rest.join(",")), "2,3");}} new C(1,2,3);
