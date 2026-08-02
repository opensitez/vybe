// vybe-test: js/map_set_iterator_more_matrix/map_constructor_duplicate_key_last_value_wins
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

const m = new Map([["a", 1], ["a", 2]]);
__check(__line(m.size), "1");
__check(__line(m.get("a")), "2");
