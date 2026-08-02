// vybe-test: js/map_set_iterator_more_matrix/map_entries_iterator_next_returns_key_value_pair
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

const pair = new Map([["a", 1]]).entries().next().value;
__check(__line(pair[0]), "a");
__check(__line(pair[1]), "1");
