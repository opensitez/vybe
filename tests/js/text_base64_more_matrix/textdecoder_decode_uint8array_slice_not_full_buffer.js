// vybe-test: js/text_base64_more_matrix/textdecoder_decode_uint8array_slice_not_full_buffer
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

const arr = new Uint8Array([88, 89, 90]);
__check(__line(new TextDecoder().decode(arr.slice(1))), "YZ");
