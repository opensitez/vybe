// vybe-test: js/es2023_2025_features/string_replace_all
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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

const s = "foo bar foo baz foo";
__check(__line(s.replaceAll("foo", "qux")), "qux bar qux baz qux");
__check(__line("a.b.c".replaceAll(".", "-")), "a-b-c");
