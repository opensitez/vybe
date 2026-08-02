// vybe-test: js/class_super_method_property_access/test_js_class_super_access_in_arrow_function_inside_method
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
    fetch() { return "BaseData"; }
}
class Sub extends Base {
    getFetcher() {
        return () => super.fetch(); // Arrow function inherits HomeObject for super!
    }
}
const s = new Sub();
const fetcher = s.getFetcher();
console.log(fetcher());
