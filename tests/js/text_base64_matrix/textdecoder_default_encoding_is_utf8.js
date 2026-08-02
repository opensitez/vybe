// vybe-test: js/text_base64_matrix/textdecoder_default_encoding_is_utf8
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

const decoder = new TextDecoder();
__check(__line(decoder.encoding), "utf-8");
