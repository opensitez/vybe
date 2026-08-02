// vybe-test: js/map_set_deep_patterns/set_nan_deduplication
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

const s = new Set([NaN, NaN, 1, 1, undefined, undefined]);
__check(__line(s.size), "3");
__check(__line(s.has(NaN)), "true");
__check(__line(s.has(undefined)), "true");
