// vybe-test: js/class_super_method_property_access/test_js_class_super_method_generator_yield_star
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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
    *generate() {
        yield 1; yield 2;
    }
}
class Sub extends Base {
    *generate() {
        yield* super.generate();
        yield 3;
    }
}
__check(__line([...new Sub().generate()].join(",")), "1,2,3");
