// vybe-test: js/string_es2022/string_matchall_with_capture_groups
// origin: languages/js/tests/js/test_string_es2022.rs

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

const str = "2024-01-15 2024-12-31";
const matches = [...str.matchAll(/(\d{4})-(\d{2})-(\d{2})/g)];
__check(__line(matches.length), "2");
__check(__line(matches[0][1]), "2024");
