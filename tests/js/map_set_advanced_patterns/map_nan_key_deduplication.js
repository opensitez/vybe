// vybe-test: js/map_set_advanced_patterns/map_nan_key_deduplication
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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
m.set(NaN, "val1");
m.set(NaN, "val2");
__check(__line(m.size + "|" + m.get(NaN)), "1|val2");
