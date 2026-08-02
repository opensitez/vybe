// vybe-test: js/string_fundamentals/trim_start_end
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

__check(__line("  spaced  ".trimStart()), "spaced  ");
__check(__line("  spaced  ".trimEnd()), "  spaced");
const noTrim = "abc".trimStart().trimEnd();
__check(__line(noTrim), "abc");
