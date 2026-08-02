// vybe-test: js/error_prototype_methods/error_to_string_on_subclass_with_empty_name
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

class Silent extends Error {} const e=new Silent("m");e.name="";__check(__line(e.toString()), "m");
