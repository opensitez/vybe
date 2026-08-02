// vybe-test: js/map_set_get_has_add_delete_clear/test_js_map_object_and_function_keys
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

const objKey = { id: 1 };
const fnKey = function() {};
const map = new Map();
map.set(objKey, "ObjectVal");
map.set(fnKey, "FunctionVal");

__check(__line(map.get(objKey) + "|" + map.get(fnKey) + "|" + map.get({ id: 1 })), "ObjectVal|FunctionVal|undefined");
