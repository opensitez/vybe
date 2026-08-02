// vybe-test: js/text_base64_more_matrix/textdecoder_fatal_invalid_utf8_throws
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

try {
  new TextDecoder("utf-8", { fatal: true }).decode(new Uint8Array([0xE2, 0x28, 0xA1]));
  console.log("no error");
} catch (error) {
  console.log(error instanceof Error);
}
