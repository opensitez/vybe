// vybe-test: js/map_set_deep_patterns/map_iteration_order
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

const m = new Map();
m.set("z", 3);
m.set("a", 1);
m.set("m", 2);
const keys = [...m.keys()];
const vals = [...m.values()];
__check(__line(keys.join(",")), "z,a,m");
__check(__line(vals.join(",")), "3,1,2");
