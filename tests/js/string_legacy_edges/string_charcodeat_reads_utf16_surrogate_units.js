// vybe-test: js/string_legacy_edges/string_charcodeat_reads_utf16_surrogate_units
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

const s = "😀";
__check(__line(s.charCodeAt(0)), "55357");
__check(__line(s.charCodeAt(1)), "56832");
