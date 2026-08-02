// vybe-test: js/class_override_getter_setter_super/test_js_class_override_value_property_in_subclass_shadows_parent_getter
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
    get title() { return "ParentTitleGetter"; }
}
class Child extends Parent {
    constructor() {
        super();
        this.title = "ChildInstanceField";
    }
}
__check(__line(new Child().title), "ChildInstanceField");
