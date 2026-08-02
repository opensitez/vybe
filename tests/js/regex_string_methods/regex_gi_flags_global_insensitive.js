// vybe-test: js/regex_string_methods/regex_gi_flags_global_insensitive
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

const matches = "Hello HELLO hello".match(/hello/gi);
__check(__line(matches.length), "3");
__check(__line(matches.join(",")), "Hello,HELLO,hello");
