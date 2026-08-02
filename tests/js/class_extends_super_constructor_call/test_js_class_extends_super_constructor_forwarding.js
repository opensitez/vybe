// vybe-test: js/class_extends_super_constructor_call/test_js_class_extends_super_constructor_forwarding
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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
    constructor(name) {
        this.name = name;
    }
}
class Child extends Parent {
    constructor(name, age) {
        super(name);
        this.age = age;
    }
}
const c = new Child("Alice", 10);
__check(__line(`${c.name}:${c.age}`), "Alice:10");
