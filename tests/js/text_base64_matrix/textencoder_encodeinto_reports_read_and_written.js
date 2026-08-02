// vybe-test: js/text_base64_matrix/textencoder_encodeinto_reports_read_and_written
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

const encoder = new TextEncoder();
const dest = new Uint8Array(10);
const result = encoder.encodeInto("Hi", dest);
__check(__line(result.read), "2");
__check(__line(result.written), "2");
__check(__line(Array.from(dest.slice(0, 2)).join(",")), "72,105");
