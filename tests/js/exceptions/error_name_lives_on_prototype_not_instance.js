// vybe-test: js/exceptions/error_name_lives_on_prototype_not_instance
// origin: languages/js/tests/js/test_exceptions.rs

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

const e = new RangeError("x");
__check(__line(Object.prototype.hasOwnProperty.call(e, "name")), "false");
__check(__line(e.name), "RangeError");
