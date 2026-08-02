// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_value_update_does_not_change_size
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

const map = new Map();
map.set("key", 100);
__check(__line(map.size), "1");
map.set("key", 200); // Replaces existing value
__check(__line(map.get("key") + "|size=" + map.size), "200|size=1");
