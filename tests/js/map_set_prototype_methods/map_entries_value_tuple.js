// vybe-test: js/map_set_prototype_methods/map_entries_value_tuple
// origin: languages/js/tests/js/test_map_set_prototype_methods.rs

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

const e=[...new Map([["k",1]]).entries()][0]; __check(__line(Array.isArray(e)), "true"); __check(__line(e.length), "2");
