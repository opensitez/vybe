// vybe-test: js/class_fields_advanced/private_field_in_operator
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

class Tagged {
    #tag = true;
    static isTagged(obj) { return #tag in obj; }
}
const t = new Tagged();
__check(__line(Tagged.isTagged(t)), "true");
__check(__line(Tagged.isTagged({})), "false");
