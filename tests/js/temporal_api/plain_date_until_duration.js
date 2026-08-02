// vybe-test: js/temporal_api/plain_date_until_duration
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

const start = new Date(2024, 0, 1).getTime();
const end   = new Date(2024, 0, 11).getTime();
const days  = Math.round((end - start) / 86400000);
__check(__line(days), "10");
