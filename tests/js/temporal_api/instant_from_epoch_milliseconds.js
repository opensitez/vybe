// vybe-test: js/temporal_api/instant_from_epoch_milliseconds
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

const epochMs = 0;
const epochSeconds = new Date(epochMs).getTime() / 1000;
__check(__line(epochSeconds), "0");
