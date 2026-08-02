// vybe-test: js/string_is_well_formed_to_well_formed/test_js_string_to_well_formed_replaces_reversed_surrogate_pair
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

const reversedSurrogates = "\uDE00\uD83D"; // Trail surrogate followed by lead surrogate
const wellFormed = reversedSurrogates.toWellFormed();
__check(__line(wellFormed.length + "|isWellFormed=" + wellFormed.isWellFormed()), "2|isWellFormed=true");
