// vybe-test: js/string_array_advanced/string_startswith_and_endswith
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let s = "javascript";
__check(__line(s.startsWith("java")), "true");
__check(__line(s.endsWith("script")), "true");
__check(__line(s.endsWith("java")), "false");
