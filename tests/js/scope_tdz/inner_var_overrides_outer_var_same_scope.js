// vybe-test: js/scope_tdz/inner_var_overrides_outer_var_same_scope
// origin: languages/js/tests/js/test_scope_tdz.rs

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

var y = "first";
{
    var y = "second";
}
__check(__line(y), "second");
