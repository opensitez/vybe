// vybe-test: js/string_array_advanced/string_includes_with_position
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

let s = "banana";
__check(__line(s.includes("an", 1)), "true");
__check(__line(s.includes("ba", 1)), "false");
