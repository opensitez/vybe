// vybe-test: js/spread_rest_advanced/rest_params_collects_extra
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

function log(first, ...others) {
  __check(__line(first), "1");
  __check(__line(others.length), "3");
}
log(1, 2, 3, 4);
