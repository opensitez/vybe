// vybe-test: js/atomics_add_sub_and_or_xor_exchange/test_js_atomics_operations_on_non_shared_typed_array
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

const i32 = new Int32Array(1);
i32[0] = 5;
const old = Atomics.add(i32, 0, 10); // Atomics math operations work on non-shared TypedArrays as well!
__check(__line(old + "|" + i32[0]), "5|15");
