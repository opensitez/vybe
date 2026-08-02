// vybe-test: js/string_methods_more/string_match_all_requires_global
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

const str = "cat bat sat";
const matches = [...str.matchAll(/[a-z]at/g)];
__check(__line(matches.length), "3");
__check(__line(matches[0][0]), "cat");
__check(__line(matches[2][0]), "sat");
