// vybe-test: js/misc_es_features/error_cause_basic
// origin: languages/js/tests/js/test_misc_es_features.rs

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

try {
  throw new Error("outer", { cause: new Error("inner") });
} catch (e) {
  __check(__line(e.message), "outer");
  __check(__line(e.cause.message), "inner");
}
