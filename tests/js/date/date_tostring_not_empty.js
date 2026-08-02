// vybe-test: js/date/date_tostring_not_empty
// origin: languages/js/tests/js/test_date.rs

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

let d = new Date(2024, 0, 1);
let s = d.toString();
__check(__line(s.length > 0), "true");
__check(__line(typeof s), "string");
