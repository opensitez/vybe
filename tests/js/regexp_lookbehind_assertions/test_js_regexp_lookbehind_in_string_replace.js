// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_lookbehind_in_string_replace
// origin: languages/js/tests/js/test_js_regexp_lookbehind_assertions.rs

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

const re = /(?<=\$)(\d+)/g;
const res = "Item1: $10, Item2: $20, Fee: 5".replace(re, (m, val) => String(Number(val) * 2));
__check(__line(res), "Item1: $20, Item2: $40, Fee: 5");
