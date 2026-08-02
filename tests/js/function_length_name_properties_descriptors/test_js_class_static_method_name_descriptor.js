// vybe-test: js/function_length_name_properties_descriptors/test_js_class_static_method_name_descriptor
// origin: languages/js/tests/js/test_js_function_length_name_properties_descriptors.rs

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

class C { static m() {} }
const desc = Object.getOwnPropertyDescriptor(C.m, "name");
__check(__line(`${desc.writable}:${desc.enumerable}:${desc.configurable}`), "false:false:true");
