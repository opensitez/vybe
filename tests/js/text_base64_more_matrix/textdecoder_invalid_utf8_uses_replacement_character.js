// vybe-test: js/text_base64_more_matrix/textdecoder_invalid_utf8_uses_replacement_character
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

const bytes = new Uint8Array([0xE2, 0x28, 0xA1]);
__check(__line(new TextDecoder().decode(bytes).includes("\uFFFD")), "true");
