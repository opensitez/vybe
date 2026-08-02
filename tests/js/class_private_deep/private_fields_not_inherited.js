// vybe-test: js/class_private_deep/private_fields_not_inherited
// origin: languages/js/tests/js/test_class_private_deep.rs

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
    #x = 42;
    getX() { return this.#x; }
}
class Child extends Parent {
    // Cannot access Parent's #x directly
    getFromParent() { return this.getX(); }
}
const c = new Child();
console.log(c.getFromParent());
