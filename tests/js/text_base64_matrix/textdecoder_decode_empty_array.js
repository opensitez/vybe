// vybe-test: js/text_base64_matrix/textdecoder_decode_empty_array
// origin: languages/js/tests/js/test_text_base64_matrix.rs

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

__check(__line(new TextDecoder().decode(new Uint8Array()) === ""), "true");
