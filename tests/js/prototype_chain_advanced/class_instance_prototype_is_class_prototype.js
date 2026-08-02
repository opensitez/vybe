// vybe-test: js/prototype_chain_advanced/class_instance_prototype_is_class_prototype
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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
    bar() { return "bar"; }
}
const f = new Foo();
__check(__line(Object.getPrototypeOf(f) === Foo.prototype), "true");
__check(__line(f.bar()), "bar");
