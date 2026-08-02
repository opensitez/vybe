// vybe-test: js/function_metadata_constructor/anonymous_function_assigned_to_object_property_infers_property_name
// origin: languages/js/tests/js/test_function_metadata_constructor.rs

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

const obj = { method: function() {} };
__check(__line(obj.method.name), "method");
