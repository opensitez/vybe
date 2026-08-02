// vybe-test: js/string_array_advanced/array_values_iterator_sequence
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

let iter = [10, 20].values();
__check(__line(iter.next().value), "10");
__check(__line(iter.next().value), "20");
__check(__line(iter.next().done), "true");
