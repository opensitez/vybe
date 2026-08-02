// vybe-test: js/string_fundamentals/string_replace_regex_and_capture_groups
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

const input = "a1b2c3";
const replaced = input.replace(/(\d)/g, "[$1]");
__check(__line(replaced), "a[1]b[2]c[3]");
__check(__line("abc123".replace(/(ab)(c)/, "$2-$1")), "c-ab123");
