// vybe-test: js/class_super_method_property_access/test_js_class_super_method_call_in_object_literal_concise_method
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

const parent = {
    greet() { return "ParentGreet"; }
};
const child = {
    __proto__: parent,
    greet() {
        return super.greet() + "Child";
    }
};
__check(__line(child.greet()), "ParentGreetChild");
