// vybe-test: js/class_extends_super_constructor_call/test_js_class_prototype_property_descriptor_non_enumerable
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

class Foo {}
const desc = Object.getOwnPropertyDescriptor(Foo, "prototype");
__check(__line(desc.writable + "|" + desc.enumerable + "|" + desc.configurable), "false|false|false");
