// vybe-test: js/map_set_prototype_methods/map_key_string_number_distinct
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

const m=new Map(); m.set(1,"n"); m.set("1","s"); __check(__line(m.get(1)), "n"); __check(__line(m.get("1")), "s");
