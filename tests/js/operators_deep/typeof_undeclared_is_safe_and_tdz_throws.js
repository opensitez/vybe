// vybe-test: js/operators_deep/typeof_undeclared_is_safe_and_tdz_throws
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(typeof doesNotExist), "undefined");
try {
    eval("{ console.log(typeof tdzVar); let tdzVar = 123; }");
} catch (e) {
    __check(__line("TDZ"), "TDZ");
}
