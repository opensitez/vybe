// vybe-test: js/string_array_advanced/array_keys_iterator_sequence
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

let iter = ["a", "b"].keys();
__check(__line(iter.next().value), "0");
__check(__line(iter.next().value), "1");
__check(__line(iter.next().done), "true");
