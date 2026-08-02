// vybe-test: js/globalthis_globals/undefined_is_a_global
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

__check(__line(typeof undefined), "undefined");
__check(__line(undefined === void 0), "true");
