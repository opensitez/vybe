// vybe-test: js/type_coercion_deep/addition_string_vs_number
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(1 + 2), "3");
__check(__line("1" + 2), "12");
__check(__line(1 + "2"), "12");
__check(__line("" + 1 + 2), "12");
__check(__line(1 + 2 + "3"), "33");
