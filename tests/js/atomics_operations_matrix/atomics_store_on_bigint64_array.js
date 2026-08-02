// vybe-test: js/atomics_operations_matrix/atomics_store_on_bigint64_array
// origin: languages/js/tests/js/test_atomics_operations_matrix.rs

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

const sab=new SharedArrayBuffer(8); const ia=new BigInt64Array(sab); __check(__line(Atomics.store(ia,0,11n)), "11n");
