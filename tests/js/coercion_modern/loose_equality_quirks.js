// vybe-test: js/coercion_modern/loose_equality_quirks
// origin: languages/js/tests/js/test_coercion_modern.rs

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

__check(__line(null == undefined), "true");
__check(__line(null == 0), "false");
__check(__line(null == ""), "false");
__check(__line(null == false), "false");
__check(__line("" == 0), "true");
__check(__line("" == false), "true");
__check(__line("0" == false), "true");
__check(__line([] == false), "true");
__check(__line([] == 0), "true");
