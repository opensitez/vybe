// vybe-test: js/string_unicode/normalize_nfd_returns_decomposed_form
// origin: languages/js/tests/js/test_string_unicode.rs

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

const nfd = "\u00E9".normalize("NFD"); // é split into e + combining accent
__check(__line(nfd.length), "2");
__check(__line(nfd === "e\u0301"), "true");
