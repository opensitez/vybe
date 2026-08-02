// vybe-test: js/class_static_deep/static_private_not_accessible_outside
// origin: languages/js/tests/js/test_class_static_deep.rs

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
    static #secret = 42;
    static get() { return Foo.#secret; }
}
__check(__line(Foo.get()), "42");
const key = "#" + "secret";
__check(__line(Foo[key] === undefined), "true");
