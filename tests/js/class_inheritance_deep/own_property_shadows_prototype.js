// vybe-test: js/class_inheritance_deep/own_property_shadows_prototype
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
    get value() { return "prototype"; }
}
class Child extends Base {
    constructor() {
        super();
        Object.defineProperty(this, "value", { value: "own", writable: true, configurable: true, enumerable: true });
    }
}
const c = new Child();
__check(__line(c.value), "own");
