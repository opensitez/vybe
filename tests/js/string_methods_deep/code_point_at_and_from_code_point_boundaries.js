// vybe-test: js/string_methods_deep/code_point_at_and_from_code_point_boundaries
// origin: languages/js/tests/js/test_string_methods_deep.rs

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

__check(__line("A".codePointAt(0)), "65");
const ascii = "ab";
__check(__line(ascii.codePointAt(0)), "97");
__check(__line(ascii.codePointAt(1)), "98");
__check(__line(String.fromCodePoint(0x10FFFF).length), "2");
