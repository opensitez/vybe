// vybe-test: js/text_base64_more_matrix/btoa_ascii_nul_byte_roundtrips
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

const value = "A\u0000B";
__check(__line(atob(btoa(value)).length), "3");
