// vybe-test: js/regex_string_methods/search_ignores_global_flag
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

// search always returns first match index, global flag doesn't matter
const re = /\d+/g;
re.lastIndex = 5; // should be ignored
console.log("abc123def456".search(re));
