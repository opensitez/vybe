// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_biguint64_typed_array
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

const sab = new SharedArrayBuffer(8);
const bu64 = new BigUint64Array(sab);
bu64[0] = 200n;
const old = Atomics.sub(bu64, 0, 50n);
__check(__line(old.toString() + "|" + bu64[0].toString()), "200|150");
