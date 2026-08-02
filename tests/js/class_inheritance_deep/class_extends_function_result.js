// vybe-test: js/class_inheritance_deep/class_extends_function_result
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

function mixin(Base) {
    return class extends Base {
        hello() { return "hello from mixin"; }
    };
}
class Foo {}
class Bar extends mixin(Foo) {}
const b = new Bar();
__check(__line(b instanceof Foo), "true");
__check(__line(b.hello()), "hello from mixin");
