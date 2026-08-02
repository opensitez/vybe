// vybe-test: js/class_private_advanced/private_field_in_subclass_separate_from_parent
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    #secret = "parent-secret";
    getSecret() { return this.#secret; }
}
class Child extends Parent {
    #secret = "child-secret";
    getChildSecret() { return this.#secret; }
}
const c = new Child();
__check(__line(c.getSecret()), "parent-secret");
__check(__line(c.getChildSecret()), "child-secret");
