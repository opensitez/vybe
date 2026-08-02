// vybe-test: js/string_array_advanced/array_of_single_number_differs_from_constructor
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

__check(__line(Array.of(3).length), "1");
__check(__line(new Array(3).length), "3");
__check(__line(Array.of(3)[0]), "3");
