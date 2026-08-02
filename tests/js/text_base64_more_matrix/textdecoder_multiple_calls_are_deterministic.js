// vybe-test: js/text_base64_more_matrix/textdecoder_multiple_calls_are_deterministic
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

const dec = new TextDecoder();
__check(__line(dec.decode(new Uint8Array([79, 75])) === dec.decode(new Uint8Array([79, 75]))), "true");
