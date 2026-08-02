// vybe-test: js/template_literal_advanced/template_with_null_and_undefined
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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

__check(__line(`${null}`), "null");
__check(__line(`${undefined}`), "undefined");
__check(__line(`${false}`), "false");
__check(__line(`${0}`), "0");
