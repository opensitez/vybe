// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_coerces_index_to_integer
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

const i32 = new Int32Array(new SharedArrayBuffer(8));
i32[1] = 10;
Atomics.add(i32, "1.9", 5); // Coerces "1.9" to index 1
__check(__line(i32[1]), "15");
