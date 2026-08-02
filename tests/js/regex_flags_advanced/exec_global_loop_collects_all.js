// vybe-test: js/regex_flags_advanced/exec_global_loop_collects_all
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /(\w+)=(\d+)/g;
const str = "a=1 b=2 c=3";
const matches = [];
let m;
while ((m = re.exec(str)) !== null) {
    matches.push(m[1] + ":" + m[2]);
}
console.log(matches.join(","));
