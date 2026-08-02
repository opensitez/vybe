// vybe-test: js/text_base64_more_matrix/textencoder_encodeinto_partial_buffer_reports_partial_write
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

const encoder = new TextEncoder();
const dest = new Uint8Array(2);
const result = encoder.encodeInto("hello", dest);
__check(__line(result.read), "2");
__check(__line(result.written), "2");
