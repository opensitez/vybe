// vybe-test: js/temporal_api/plain_time_add
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

const startH = 10, startM = 30;
const addH = 2, addM = 15;
const totalM = startH * 60 + startM + addH * 60 + addM;
__check(__line(Math.floor(totalM / 60)), "12");
__check(__line(totalM % 60), "45");
