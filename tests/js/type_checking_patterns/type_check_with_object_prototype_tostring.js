// vybe-test: js/type_checking_patterns/type_check_with_object_prototype_tostring
// origin: languages/js/tests/js/test_type_checking_patterns.rs

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

function typeOf(val) {
    return Object.prototype.toString.call(val).slice(8, -1);
}
__check(__line(typeOf(null)), "Null");
__check(__line(typeOf(undefined)), "Undefined");
__check(__line(typeOf(42)), "Number");
__check(__line(typeOf("str")), "String");
__check(__line(typeOf([])), "Array");
__check(__line(typeOf({})), "Object");
__check(__line(typeOf(/re/)), "RegExp");
