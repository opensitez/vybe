// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_regexp_object
// origin: languages/js/tests/js/test_js_structured_clone_map_set_regexp_date.rs

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

const re = /test_\d+/giu;
const clone = structuredClone(re);
__check(__line((clone !== re) + "|" + (clone instanceof RegExp) + "|" + clone.source + "|" + clone.flags), "true|true|test_\\d+|giu");
