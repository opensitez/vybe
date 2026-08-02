// vybe-test: js/class_inheritance_advanced/static_members_inherited_and_overridden
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
    static version() { return "base"; }
    static metadata() { return "Base:" + this.name + ":" + this.version(); }
}

class Child extends Base {
    static version() { return "child"; }
    static metadataFromSuper() { return super.version() + "|" + super.metadata(); }
}

__check(__line(Object.getPrototypeOf(Child) === Base), "true");
__check(__line(Child.version()), "child");
__check(__line(Child.metadataFromSuper()), "base|Base:Child:child");
