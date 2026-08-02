// vybe-test: js/ecma_strings/string_code_points
// origin: languages/js/tests/js/test_ecma_strings.rs

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

__check(__line("hello".charCodeAt(1)), "101");
__check(__line("hello".charAt(-1)), "");
__check(__line(String.fromCharCode(72, 105, 33)), "Hi!");
