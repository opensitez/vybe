// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_or_bitwise_operation
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

const sab = new SharedArrayBuffer(4);
const i32 = new Int32Array(sab);
i32[0] = 0b0101;
const old = Atomics.or(i32, 0, 0b1010);
__check(__line(old + "|" + i32[0]), "5|15");
