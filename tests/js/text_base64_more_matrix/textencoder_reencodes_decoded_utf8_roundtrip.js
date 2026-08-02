// vybe-test: js/text_base64_more_matrix/textencoder_reencodes_decoded_utf8_roundtrip
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

const s = new TextDecoder().decode(new Uint8Array([195, 169]));
__check(__line(Array.from(new TextEncoder().encode(s)).join(",")), "195,169");
