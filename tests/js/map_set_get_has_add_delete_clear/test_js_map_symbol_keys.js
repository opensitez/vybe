// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_symbol_keys
// origin: languages/js/tests/js/test_js_map_set_get_has_add_delete_clear.rs

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

const s1 = Symbol("s1");
const s2 = Symbol("s2");
const map = new Map();
map.set(s1, "Val1");
map.set(s2, "Val2");

__check(__line(map.get(s1) + "|" + map.get(s2)), "Val1|Val2");
