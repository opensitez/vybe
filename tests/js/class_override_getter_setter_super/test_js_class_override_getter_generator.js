// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_generator
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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
    get stream() {
        return function*() { yield 1; yield 2; };
    }
}
class Derived extends Base {
    get stream() {
        const baseStream = super.stream;
        return function*() {
            yield* baseStream();
            yield 3;
        };
    }
}
__check(__line([...new Derived().stream()].join(",")), "1,2,3");
