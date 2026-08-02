// vybe-test: js/class_super_method_property_access/test_js_class_super_property_delete_throws_referenceerror
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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

class Base {
    foo() {}
}
class Sub extends Base {
    deleteFoo() {
        try {
            eval("delete super.foo;");
        } catch (e) {
            __check(__line("Delete Super ReferenceError"), "Delete Super ReferenceError");
        }
    }
}
new Sub().deleteFoo();
