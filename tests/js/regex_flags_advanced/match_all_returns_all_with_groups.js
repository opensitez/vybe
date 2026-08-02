// vybe-test: js/regex_flags_advanced/match_all_returns_all_with_groups
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

const re = /(\w+)=(\d+)/g;
const results = [...("a=1 b=2".matchAll(re))];
__check(__line(results.length), "2");
__check(__line(results[0][1]), "a");
__check(__line(results[1][2]), "2");
