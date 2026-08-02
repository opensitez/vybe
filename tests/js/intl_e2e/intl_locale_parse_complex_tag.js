// vybe-test: js/intl_e2e/intl_locale_parse_complex_tag
// origin: languages/js/tests/js/test_intl_e2e.rs

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

const loc = new Intl.Locale("zh-Hans-CN");
        __check(__line(loc.language, loc.script, loc.region), "zh Hans CN");
