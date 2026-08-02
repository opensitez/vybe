// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_side_effect_count
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

let reads = 0;
class Base {
    get count() { reads++; return 1; }
}
class Derived extends Base {
    get count() { return super.count + super.count; }
}
__check(__line(new Derived().count + "|Reads=" + reads), "2|Reads=2");
