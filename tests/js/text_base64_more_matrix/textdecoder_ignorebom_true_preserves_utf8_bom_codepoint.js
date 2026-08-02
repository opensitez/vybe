// vybe-test: js/text_base64_more_matrix/textdecoder_ignorebom_true_preserves_utf8_bom_codepoint
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

const bytes = new Uint8Array([239, 187, 191, 65]);
const value = new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
__check(__line(value.charCodeAt(0)), "65279");
__check(__line(value.slice(1)), "A");
