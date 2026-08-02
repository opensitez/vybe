// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_sub_underflow_wrap_around
// origin: languages/js/tests/js/test_js_atomics_add_sub_and_or_xor_exchange.rs

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

const u8 = new Uint8Array(new SharedArrayBuffer(1));
u8[0] = 0;
Atomics.sub(u8, 0, 1);
__check(__line(u8[0]), "255");
