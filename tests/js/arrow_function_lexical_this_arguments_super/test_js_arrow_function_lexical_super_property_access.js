// vybe-test: js/arrow_function_lexical_this_arguments_super/test_js_arrow_function_lexical_super_property_access
// origin: languages/js/tests/js/test_js_arrow_function_lexical_this_arguments_super.rs

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
    greet() { return "BaseGreet"; }
}
class Derived extends Base {
    getArrowGreet() {
        return () => super.greet();
    }
}
const d = new Derived();
__check(__line(d.getArrowGreet()()), "BaseGreet");
