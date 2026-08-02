// vybe-test: js/map_set_deep/map_can_use_any_value_as_key
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

const m = new Map();
m.set(null, "null key");
m.set(undefined, "undefined key");
m.set(NaN, "nan key");
m.set(true, "bool key");
__check(__line(m.get(null)), "null key");
__check(__line(m.get(undefined)), "undefined key");
__check(__line(m.get(NaN)), "nan key"); // NaN === NaN in Map (SameValueZero)
__check(__line(m.get(true)), "bool key");
