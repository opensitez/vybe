// vybe-test: js/string_methods_more/string_slice_vs_substring
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
// slice: negative means from end
__check(__line(s.slice(-5)), "world");
// substring: negative treated as 0
__check(__line(s.substring(-5, 5)), "hello");
