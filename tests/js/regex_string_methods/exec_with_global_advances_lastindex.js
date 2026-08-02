// vybe-test: js/regex_string_methods/exec_with_global_advances_lastindex
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const re = /\d+/g;
const str = "1 2 3";
const matches = [];
let m;
while ((m = re.exec(str)) !== null) {
    matches.push(m[0]);
}
console.log(matches.join(","));
