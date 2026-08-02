// vybe-test: js/string_array_advanced/array_entries_iterator_exposes_index_value_pairs
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let first = ["x", "y"].entries().next().value;
__check(__line(first[0]), "0");
__check(__line(first[1]), "x");
