// vybe-test: js/string_fundamentals/normalize_nfc_composes
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

const decomposed = "e\u0301"; // e + combining acute
const composed = decomposed.normalize("NFC");
__check(__line(composed.length), "1");
__check(__line(composed === "\u00E9"), "true"); // é
