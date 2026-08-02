// vybe-test: js/misc_es_features/array_from_with_length_generates_sequence
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const squares = Array.from({ length: 5 }, (_, i) => (i + 1) ** 2);
__check(__line(squares.join(",")), "1,4,9,16,25");
