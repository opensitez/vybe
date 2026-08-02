// vybe-test: js/string_regex_integration/replace_with_dollar_signs
// origin: languages/js/tests/js/test_string_regex_integration.rs

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

// $1 = first capture group, $& = whole match
const result = "hello world".replace(/(\w+)\s(\w+)/, "$2 $1");
__check(__line(result), "world hello");
const withFull = "abc".replace(/b/, "[$&]");
__check(__line(withFull), "a[b]c");
