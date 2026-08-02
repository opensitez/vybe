// vybe-test: js/error_subclasses_eval_range_reference_syntax_type_uri/test_js_error_subclasses_prototype_chain_verification
// origin: languages/js/tests/js/test_js_error_subclasses_eval_range_reference_syntax_type_uri.rs

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

const err = new RangeError("Range");
__check(__line(Object.getPrototypeOf(RangeError.prototype) === Error.prototype), "true");
