// vybe-test: js/class_extends_super_constructor_call/test_js_class_constructor_name_property
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class NamedClass {}
__check(__line(NamedClass.name + "|" + new NamedClass().constructor.name), "NamedClass|NamedClass");
