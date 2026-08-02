// vybe-test: js/string_methods_deep/matchall_returns_all_matches
// origin: languages/js/tests/js/test_string_methods_deep.rs

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

const str = "test1 test2 test3";
const matches = [...str.matchAll(/test(\d)/g)];
__check(__line(matches.length), "3");
__check(__line(matches[0][1]), "1");
__check(__line(matches[2][1]), "3");
