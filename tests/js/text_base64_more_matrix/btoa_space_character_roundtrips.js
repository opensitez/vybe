// vybe-test: js/text_base64_more_matrix/btoa_space_character_roundtrips
// origin: languages/js/tests/js/test_text_base64_more_matrix.rs

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

__check(__line(atob(btoa(" ")) === " "), "true");
