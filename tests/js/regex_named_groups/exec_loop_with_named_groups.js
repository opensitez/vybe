// vybe-test: js/regex_named_groups/exec_loop_with_named_groups
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

const re = /(?<n>\d+)/g;
const results = [];
let m;
while ((m = re.exec("a1b22c333")) !== null) {
    results.push(m.groups.n);
}
console.log(results.join(","));
