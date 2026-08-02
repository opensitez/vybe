// vybe-test: js/map_set_iterator_more_matrix/set_add_same_object_twice_keeps_single_entry
// origin: languages/js/tests/js/test_map_set_iterator_more_matrix.rs

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

const obj = {};
const s = new Set([obj]);
s.add(obj);
__check(__line(s.size), "1");
