// vybe-test: js/class_fields_advanced/field_initializers_run_in_order_and_access_this
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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

class Sequence {
    first = 1;
    second = this.first + 1;
}
const s = new Sequence();
__check(__line(s.first), "1");
__check(__line(s.second), "2");
