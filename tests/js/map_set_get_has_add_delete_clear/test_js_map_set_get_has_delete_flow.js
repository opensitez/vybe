// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_set_get_has_delete_flow
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
map.set("key1", "val1");
map.set("key2", "val2");

__check(__line(map.get("key1") + "|" + map.has("key2") + "|size=" + map.size), "val1|true|size=2");
map.delete("key1");
__check(__line(map.has("key1") + "|size=" + map.size), "false|size=1");
