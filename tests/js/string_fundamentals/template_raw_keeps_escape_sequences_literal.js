// vybe-test: js/string_fundamentals/template_raw_keeps_escape_sequences_literal
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

const raw = String.raw`line1\nline2`;
__check(__line(raw.includes("\\n")), "true");
__check(__line(raw.split("\\n")[0]), "line1");
__check(__line(raw.split("\\n")[1]), "line2");
