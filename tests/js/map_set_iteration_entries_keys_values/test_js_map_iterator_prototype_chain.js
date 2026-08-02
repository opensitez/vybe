// vybe-test: js/map_set_iteration_entries_keys_values/test_js_map_iterator_prototype_chain
// origin: languages/js/tests/js/test_js_map_set_iteration_entries_keys_values.rs

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
const iter = map.keys();
const proto = Object.getPrototypeOf(iter);
__check(__line(typeof iter[Symbol.iterator] === "function"), "true");
