// vybe-test: js/date/date_toisostring
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

let d = new Date("2024-03-15T12:00:00Z");
let iso = d.toISOString();
__check(__line(iso.startsWith("2024-03-15")), "true");
__check(__line(iso.endsWith("Z")), "true");
