// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_lookbehind_exec_null_on_failed_assertion
// origin: languages/js/tests/js/test_js_regexp_lookbehind_assertions.rs

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

const re = /(?<=\$)(\d+)/;
__check(__line(re.exec("100 USD") === null), "true");
