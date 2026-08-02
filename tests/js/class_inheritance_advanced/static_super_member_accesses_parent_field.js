// vybe-test: js/class_inheritance_advanced/static_super_member_accesses_parent_field
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
    static level = "base";
}

class Child extends Base {
    static level = "child";
    static describe() {
        return `${super.level}:${this.level}`;
    }
}
__check(__line(Child.describe()), "base:child");
