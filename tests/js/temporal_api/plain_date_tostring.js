// vybe-test: js/temporal_api/plain_date_tostring
// origin: languages/js/tests/js/test_temporal_api.rs

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

function pad(n) { return String(n).padStart(2, "0"); }
const d = new Date(2024, 2, 5); // 2024-03-05
const s = d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
__check(__line(s), "2024-03-05");
