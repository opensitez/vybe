// vybe-test: js/map_set_get_has_add_delete_clear/test_js_set_add_has_delete_flow
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

const set = new Set();
set.add(10);
set.add(20);
set.add(10); // Duplicate ignored

__check(__line(set.size + "|" + set.has(10)), "2|true");
set.delete(10);
__check(__line(set.size + "|" + set.has(10)), "1|false");
