// vybe-test: js/spread_rest_deep/spread_converts_map_to_entries
// origin: languages/js/tests/js/test_spread_rest_deep.rs

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

const map = new Map([["a", 1], ["b", 2]]);
const entries = [...map];
__check(__line(entries.map(([k,v]) => k+"="+v).join(",")), "a=1,b=2");
