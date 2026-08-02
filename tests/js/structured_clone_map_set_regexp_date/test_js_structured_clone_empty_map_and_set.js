// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_empty_map_and_set
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

const m = new Map();
const s = new Set();
const cm = structuredClone(m);
const cs = structuredClone(s);
__check(__line(cm.size + "|" + cs.size), "0|0");
