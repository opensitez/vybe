// vybe-test: js/text_base64_matrix/textdecoder_decode_from_arraybuffer
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

const bytes = new Uint8Array([79, 75]);
__check(__line(new TextDecoder().decode(bytes.buffer)), "OK");
