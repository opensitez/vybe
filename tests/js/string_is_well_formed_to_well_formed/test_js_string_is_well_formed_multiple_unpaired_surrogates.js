// vybe-test: js/string_is_well_formed_to_well_formed/test_js_string_is_well_formed_multiple_unpaired_surrogates
// origin: languages/js/tests/js/test_js_string_is_well_formed_to_well_formed.rs

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

const multipleLone = "\uD800\uD800\uD800";
__check(__line(multipleLone.isWellFormed()), "false");
