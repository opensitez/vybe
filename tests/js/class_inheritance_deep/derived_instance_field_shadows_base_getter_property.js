// vybe-test: js/class_inheritance_deep/derived_instance_field_shadows_base_getter_property
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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
    constructor() {
        this.mode = "base-mode";
    }
}
class Child extends Base {
    constructor() {
        super();
        this.mode = "field-mode";
    }
}
const child = new Child();
__check(__line(`${child.mode}|${Object.hasOwn(child, "mode")}|${child.mode === "field-mode"}`), "field-mode|true|true");
