// vybe-test: js/string_methods_deep/padstart_with_multi_char_fill
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

__check(__line("5".padStart(5, "0")), "00005");
__check(__line("abc".padStart(7, "xy")), "xyxyabc");
