// vybe-test: js/class_inheritance_advanced/new_target_is_subclass_in_super_constructor
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
        this.constructedAs = this.constructor.name;
    }
}
class Derived extends Base {}
const b = new Base();
const d = new Derived();
__check(__line(b.constructedAs), "Base");
__check(__line(d.constructedAs), "Derived");
