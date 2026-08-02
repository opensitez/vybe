// vybe-test: js/string_es2022/string_matchall_indices
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

const matches = [...("abcabc".matchAll(/a/g))];
const indices = matches.map(m => m.index).join(",");
__check(__line(indices), "0,3");
