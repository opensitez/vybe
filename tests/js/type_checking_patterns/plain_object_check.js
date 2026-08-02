// vybe-test: js/type_checking_patterns/plain_object_check
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

function isPlainObject(val) {
    if (typeof val !== "object" || val === null) return false;
    const proto = Object.getPrototypeOf(val);
    return proto === Object.prototype || proto === null;
}
__check(__line(isPlainObject({})), "true");
__check(__line(isPlainObject(Object.create(null))), "true");
__check(__line(isPlainObject([])), "false");
__check(__line(isPlainObject(new Date())), "false");
__check(__line(isPlainObject(42)), "false");
