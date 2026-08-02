// vybe-test: js/string_legacy_edges/string_normalize_nfd_expands_composed_character
// origin: languages/js/tests/js/test_string_legacy_edges.rs

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

const s = "é".normalize("NFD");
__check(__line(s.length), "2");
__check(__line(s === "e\u0301"), "true");
