// vybe-test: js/prototype_patterns_deep/constructor_property_preserved
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

class Foo {}
const f = new Foo();
__check(__line(f.constructor === Foo), "true");
__check(__line(f.constructor.name), "Foo");
