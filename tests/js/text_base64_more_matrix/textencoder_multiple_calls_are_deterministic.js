// vybe-test: js/text_base64_more_matrix/textencoder_multiple_calls_are_deterministic
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

const enc = new TextEncoder();
__check(__line(Array.from(enc.encode("ok")).join(",") === Array.from(enc.encode("ok")).join(",")), "true");
