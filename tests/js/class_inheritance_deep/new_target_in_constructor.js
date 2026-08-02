// vybe-test: js/class_inheritance_deep/new_target_in_constructor
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

class Foo {
    constructor() {
        this.target = this.constructor.name;
    }
}
class Bar extends Foo {}
const f = new Foo();
const b = new Bar();
__check(__line(f.target), "Foo");
__check(__line(b.target), "Bar");
