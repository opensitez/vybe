// vybe-test: js/string_methods_more/string_replace_all_and_includes_positioning
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

__check(__line("foo bar foo".replaceAll("foo", "baz")), "baz bar baz");
__check(__line("hello world".includes("wor")), "true");
__check(__line("hello world".includes("world", 10)), "false");
