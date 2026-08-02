// vybe-test: js/class_inheritance_advanced/new_target_reflects_constructed_class_in_inheritance_chain
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
    constructor() {
        this.requested = new.target;
    }
}
class Child extends Base {}
class GrandChild extends Child {}
__check(__line(new Child().requested.name), "Child");
__check(__line(new GrandChild().requested.name), "GrandChild");
