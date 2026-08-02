// vybe-test: js/string_array_advanced/string_charcodeat_and_fromcharcode_roundtrip
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

let code = "A".charCodeAt(0);
__check(__line(code), "65");
__check(__line(String.fromCharCode(code + 1)), "B");
