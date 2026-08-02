// vybe-test: js/class_inheritance_advanced/static_and_instance_properties_show_expected_resolution_order
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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
    static namespace = "base";
    static get label() { return `static:${this.namespace}`; }
    marker = "base";
    label() { return "base"; }
}

class Child extends Base {
    static namespace = "child";
    static get fullLabel() {
        return `${super.namespace}|${this.namespace}|${super.label}`;
    }
    marker = "child";
    label() { return `child:${super.label()}`; }
}

const c = new Child();
__check(__line(c.marker), "child");
__check(__line(c.label()), "child:base");
__check(__line(Child.namespace), "child");
__check(__line(Child.fullLabel), "base|child|static:child");
__check(__line(Child.label), "static:child"); // method name, not executed
