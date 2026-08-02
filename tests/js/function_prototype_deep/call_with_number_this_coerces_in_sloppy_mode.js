// vybe-test: js/function_prototype_deep/call_with_number_this_coerces_in_sloppy_mode
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function tag() { return Object.prototype.toString.call(this); } __check(__line(tag.call(42).includes("Number")), "true");
