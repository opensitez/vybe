// vybe-test: js/string_methods_more/string_split_empty_separator
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

const chars = "abc".split("");
__check(__line(chars.length), "3");
__check(__line(chars.join("-")), "a-b-c");
