// vybe-test: js/globalthis_globals/global_isfinite_coerces_argument
// origin: languages/js/tests/js/test_globalthis_globals.rs

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

__check(__line(isFinite("5")), "true");
__check(__line(isFinite(null)), "true");
__check(__line(isFinite(Infinity)), "false");
__check(__line(isFinite(NaN)), "false");
