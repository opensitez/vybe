// vybe-test: js/string_methods_more/string_starts_ends_with_position_from_index
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

const s = "hello world";
__check(__line(s.startsWith("world", 6)), "true");
__check(__line(s.endsWith("hello", 5)), "true");
