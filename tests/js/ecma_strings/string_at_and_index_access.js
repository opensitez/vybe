// vybe-test: js/ecma_strings/string_at_and_index_access
// origin: languages/js/tests/js/test_ecma_strings.rs

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

__check(__line("abc".at(1)), "b");
__check(__line("abc".at(-1)), "c");
__check(__line("".at(0)), "undefined");
__check(__line("abc"[2]), "c");
