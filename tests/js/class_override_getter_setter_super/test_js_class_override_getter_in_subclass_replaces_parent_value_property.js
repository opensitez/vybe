// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_in_subclass_replaces_parent_value_property
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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

class Parent {
    title = "ParentTitle";
}
class Child extends Parent {
    get title() { return "ChildGetterTitle"; }
}
__check(__line(new Child().title), "ParentTitle");
