// vybe-test: js/temporal_api/plain_date_from_string
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

const parts = "2024-03-21".split("-").map(Number);
const d = new Date(parts[0], parts[1] - 1, parts[2]);
__check(__line(d.getFullYear()), "2024");
__check(__line(d.getMonth() + 1), "3");
__check(__line(d.getDate()), "21");
