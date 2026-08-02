// vybe-test: js/string_processing_patterns/string_lines_split_and_process
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

const text = "line1\nline2\nline3\n";
const lines = text.split("\n").filter(Boolean);
__check(__line(lines.length), "3");
__check(__line(lines[0]), "line1");
__check(__line(lines[2]), "line3");
