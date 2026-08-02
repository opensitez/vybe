// vybe-test: js/text_base64_matrix/textdecoder_decode_bom_prefixed_utf8
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

const bytes = new Uint8Array([239, 187, 191, 65]);
__check(__line(new TextDecoder().decode(bytes)), "A");
