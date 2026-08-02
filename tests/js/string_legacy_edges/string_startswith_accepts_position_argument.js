// vybe-test: js/string_legacy_edges/string_startswith_accepts_position_argument
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

const s = "banana";
__check(__line(s.startsWith("na", 2)), "true");
__check(__line(s.startsWith("ba", 2)), "false");
