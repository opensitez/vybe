// vybe-test: js/text_base64_matrix/textdecoder_decode_utf8_astral_bytes
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

const bytes = new Uint8Array([240, 159, 152, 128]);
const value = new TextDecoder().decode(bytes);
__check(__line(value), "😀");
__check(__line(value.length), "2");
