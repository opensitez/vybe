// vybe-test: js/object_spread_edge/spread_string_adds_indexed_properties
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

const chars = { ..."hello" };
__check(__line(chars[0]), "h");
__check(__line(chars[1]), "e");
__check(__line(chars[4]), "o");
__check(__line(Object.hasOwn(chars, "length")), "false");
