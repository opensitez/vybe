// vybe-test: js/regex_advanced/regex_exec_global_loop
// origin: languages/js/tests/js/test_regex_advanced.rs

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

let re = /\d+/g;
let s = "a1b22c333";
let results = [];
let m;
while ((m = re.exec(s)) !== null) {
    results.push(m[0]);
}
console.log(results.join(","));
