// vybe-test: js/string_fundamentals/escape_sequence_unicode_code_point
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

__check(__line("\u0041"), "A");   // A
__check(__line("\u{1F600}"), "😀"); // 😀
__check(__line("\u{0041}"), "A"); // A via curly syntax
