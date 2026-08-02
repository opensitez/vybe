// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_constructor_with_iterable_tuples
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

const entries = [["k1", 10], ["k2", 20]];
const map = new Map(entries);
__check(__line(map.get("k1") + "|" + map.get("k2")), "10|20");
