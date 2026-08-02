// vybe-test: js/string_legacy_edges/string_split_with_undefined_separator_returns_whole_string
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

const parts = "a,b,c".split(undefined);
__check(__line(parts.length), "1");
__check(__line(parts[0]), "a,b,c");
