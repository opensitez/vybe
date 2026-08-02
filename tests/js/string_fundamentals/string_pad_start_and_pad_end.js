// vybe-test: js/string_fundamentals/string_pad_start_and_pad_end
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

__check(__line("x".padStart(4, "0")), "000x");
__check(__line("x".padEnd(4, "0")), "x000");
__check(__line("x".padStart(2, "ab")), "ax");
