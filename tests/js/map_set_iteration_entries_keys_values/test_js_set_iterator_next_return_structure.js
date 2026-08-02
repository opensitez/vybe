// vybe-test: js/map_set_iteration_entries_keys_values/test_js_set_iterator_next_return_structure
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

const set = new Set([42]);
const iter = set.values();
const step1 = iter.next();
const step2 = iter.next();

__check(__line(`${step1.value}|done=${step1.done}`), "42|done=false");
__check(__line(`${step2.value}|done=${step2.done}`), "undefined|done=true");
