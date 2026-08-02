// vybe-test: js/regex_named_groups/exec_loop_with_global_flag
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

const re = /\d+/g;
const text = "a1b22c333";
const results = [];
let m;
while ((m = re.exec(text)) !== null) {
    results.push(m[0]);
}
console.log(results.join(","));
