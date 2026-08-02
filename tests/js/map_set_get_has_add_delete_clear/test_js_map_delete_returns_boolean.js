// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_delete_returns_boolean
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

const map = new Map([["exist", 1]]);
__check(__line(map.delete("exist") + "|" + map.delete("missing")), "true|false");
