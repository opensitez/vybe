// vybe-test: js/regex_flags_advanced/sticky_and_global_differ
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const sticky = /\w+/y;
const global = /\w+/g;
const str = "  foo bar";
// sticky: must match at lastIndex (0), space is not \w
const ms = sticky.exec(str);
__check(__line(ms), "null"); // null, no match at pos 0
// global: searches anywhere
const mg = global.exec(str);
__check(__line(mg[0]), "foo"); // "foo"
