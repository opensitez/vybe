// vybe-test: js/regex_named_groups/negative_lookbehind_excludes_prefix
// origin: languages/js/tests/js/test_regex_named_groups.rs

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

const re = /(?<!\$)\d+/;
const m = "price: 100, discount: $20".match(re);
__check(__line(m[0]), "100");
