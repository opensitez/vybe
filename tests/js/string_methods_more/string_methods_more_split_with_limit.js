// vybe-test: js/string_methods_more/string_methods_more_split_with_limit
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

const parts = "a,b,c".split(",", 2);
__check(__line(parts.length), "2");
__check(__line(parts.join("|")), "a|b");
