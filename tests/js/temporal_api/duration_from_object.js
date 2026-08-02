// vybe-test: js/temporal_api/duration_from_object
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

function makeDuration(obj) {
    return { years:0, months:0, days:0, hours:0, minutes:0, seconds:0, ...obj };
}
const dur = makeDuration({ days: 7, hours: 12 });
__check(__line(dur.days), "7");
__check(__line(dur.hours), "12");
