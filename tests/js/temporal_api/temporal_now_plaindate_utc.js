// vybe-test: js/temporal_api/temporal_now_plaindate_utc
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

const year = new Date().getFullYear();
__check(__line(typeof year === "number"), "true");
__check(__line(year >= 2024), "true");
