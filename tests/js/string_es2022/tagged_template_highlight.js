// vybe-test: js/string_es2022/tagged_template_highlight
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

function highlight(strings, ...vals) {
  return strings.reduce((acc, str, i) => acc + str + (vals[i] !== undefined ? "[" + vals[i] + "]" : ""), "");
}
const a = 1, b = 2;
__check(__line(highlight`sum of ${a} and ${b} is ${a + b}`), "sum of [1] and [2] is [3]");
