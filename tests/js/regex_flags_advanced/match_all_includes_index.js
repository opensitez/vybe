// vybe-test: js/regex_flags_advanced/match_all_includes_index
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

const re = /cat/g;
const matches = [...("catfish cat caterpillar".matchAll(re))];
__check(__line(matches.map(m => m.index).join(",")), "0,8,12");
