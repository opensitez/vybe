// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_shadowed_private_name_in_subclass
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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
    #tag = "Base";
    static isBase(o) { return #tag in o; }
}
class Sub extends Base {
    #tag = "Sub";
    static isSub(o) { return #tag in o; }
}
const b = new Base();
const s = new Sub();
__check(__line(`${Base.isBase(b)}|${Base.isBase(s)}|${Sub.isSub(b)}|${Sub.isSub(s)}`), "true|true|false|true");
