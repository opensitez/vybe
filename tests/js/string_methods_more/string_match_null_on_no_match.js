// vybe-test: js/string_methods_more/string_match_null_on_no_match
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

const result = "hello".match(/xyz/);
__check(__line(result), "null");
