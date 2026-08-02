// vybe-test: js/map_set_get_has_add_delete_clear/test_js_set_object_reference_uniqueness
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
const o1 = { a: 1 };
const o2 = { a: 1 };
set.add(o1);
set.add(o2);
set.add(o1); // Duplicate reference ignored

__check(__line(set.size), "2");
