// vybe-test: js/template_literal_advanced/string_raw_vs_cooked
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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

const raw = String.raw`\n\t`;
const cooked = `\n\t`;
__check(__line(raw.length), "4");   // 4 chars: \, n, \, t
__check(__line(cooked.length), "2"); // 2 chars: newline, tab
