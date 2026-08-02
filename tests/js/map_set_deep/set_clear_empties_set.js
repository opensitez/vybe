// vybe-test: js/map_set_deep/set_clear_empties_set
// origin: languages/js/tests/js/test_map_set_deep.rs

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

const s = new Set([1, 2, 3]);
s.clear();
__check(__line(s.size), "0");
__check(__line(s.has(1)), "false");
