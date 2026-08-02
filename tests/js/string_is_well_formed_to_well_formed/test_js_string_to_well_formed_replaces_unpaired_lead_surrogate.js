// vybe-test: js/string_is_well_formed_to_well_formed/test_js_string_to_well_formed_replaces_unpaired_lead_surrogate
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

const loneLead = "a\uD83Db";
const wellFormed = loneLead.toWellFormed();
__check(__line(wellFormed + "|code=" + wellFormed.charCodeAt(1)), "ab|code=65533"); // Lone surrogate replaced by U+FFFD (65533 replacement character)!
