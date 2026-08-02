// vybe-test: js/error_handling_advanced/error_type_discrimination
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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

function classify(fn) {
    try {
        fn();
    } catch(e) {
        if (e instanceof TypeError) return "type";
        if (e instanceof RangeError) return "range";
        if (e instanceof SyntaxError) return "syntax";
        return "other:" + e.constructor.name;
    }
}
__check(__line(classify(() => null.x)), "type");
__check(__line(classify(() => new Array(-1))), "range");
__check(__line(classify(() => { throw new RangeError("oops"); })), "range");
