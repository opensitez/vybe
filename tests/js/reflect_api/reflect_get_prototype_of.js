// vybe-test: js/reflect_api/reflect_get_prototype_of
// origin: languages/js/tests/js/test_reflect_api.rs

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
__check(__line(Reflect.getPrototypeOf(f) === Foo.prototype), "true");
__check(__line(Reflect.getPrototypeOf(Foo.prototype) === Object.prototype), "true");
