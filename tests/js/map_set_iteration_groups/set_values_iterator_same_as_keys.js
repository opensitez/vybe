// vybe-test: js/map_set_iteration_groups/set_values_iterator_same_as_keys
// origin: languages/js/tests/js/test_map_set_iteration_groups.rs

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

const s=new Set([1]); __check(__line(s.values().next().value), "1");__check(__line(s.keys().next().value), "1");
