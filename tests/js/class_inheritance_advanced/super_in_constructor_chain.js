// vybe-test: js/class_inheritance_advanced/super_in_constructor_chain
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

class A {
    constructor(x) { this.x = x; }
}

class B extends A {
    constructor(x, y) {
        super(x);
        this.y = y;
    }
}
class C extends B {
    constructor(x, y, z) {
        super(x, y);
        this.z = z;
    }
}
const c = new C(1, 2, 3);
__check(__line(c.x), "1");
__check(__line(c.y), "2");
__check(__line(c.z), "3");
