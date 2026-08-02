// vybe-test: js/regex_advanced/regex_exec_basic
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

let re = /(\d+)-(\d+)/;
let m = re.exec("date: 2024-01");
__check(__line(m[0]), "2024-01");
__check(__line(m[1]), "2024");
__check(__line(m[2]), "01");
