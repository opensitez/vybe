// vybe-test: js/string_is_well_formed_to_well_formed/test_js_string_to_well_formed_replaces_unpaired_trail_surrogate
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

const loneTrail = "a\uDE00b";
const wellFormed = loneTrail.toWellFormed();
__check(__line(wellFormed + "|code=" + wellFormed.charCodeAt(1)), "ab|code=65533");
