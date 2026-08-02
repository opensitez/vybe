// vybe-test: js/ecma_strings/string_iterable_iteration
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

__check(__line(Array.from("ab\u0301").join("|")), "a|b|́");
__check(__line([...new Set("abb")].join("|")), "a|b");
