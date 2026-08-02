// vybe-test: js/string_methods_more/string_code_point_round_trip
// origin: languages/js/tests/js/test_string_methods_more.rs

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

__check(__line("ABC".codePointAt(0)), "65");
__check(__line("ABC".codePointAt(1)), "66");
__check(__line("😀".length), "2");
__check(__line("😀".codePointAt(0)), "128512");
__check(__line(String.fromCodePoint(0x41, 0x42)), "AB");
