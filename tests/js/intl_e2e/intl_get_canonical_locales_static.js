// vybe-test: js/intl_e2e/intl_get_canonical_locales_static
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

const tags = Intl.getCanonicalLocales(["EN-us", "FR-fr"]);
        __check(__line(tags[0], tags[1]), "en-US fr-FR");
