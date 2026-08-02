// vybe-test: js/map_set_deep_patterns/map_chaining_pattern
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

const freq = ["a","b","a","c","b","a"].reduce((m, v) => m.set(v, (m.get(v)??0)+1), new Map());
const sorted = [...freq.entries()].sort((a,b) => b[1]-a[1]);
__check(__line(sorted[0].join("=")), "a=3");
__check(__line(sorted[1].join("=")), "b=2");
