// vybe-test: js/object_spread_edge/spread_overrides_with_explicit_property
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const defaults = { color: "red", size: "M" };
const custom = { ...defaults, color: "blue" };
__check(__line(custom.color), "blue");
__check(__line(custom.size), "M");
