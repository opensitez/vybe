// vybe-test: js/regex_named_groups/lookahead_negative_excludes_suffix
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

const matches = [..."12px 30em 5px".matchAll(/\d+(?!px)\b/g)].map(m => m[0]);
__check(__line(matches.join(",")), "");
