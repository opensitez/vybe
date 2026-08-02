// vybe-test: js/globalthis_globals/global_encode_decode_uri_components
// origin: languages/js/tests/js/test_globalthis_globals.rs

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

const enc = encodeURIComponent("hello world");
__check(__line(enc), "hello%20world");
__check(__line(decodeURIComponent(enc)), "hello world");
