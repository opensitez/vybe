// vybe-test: js/atomics_operations_matrix/atomics_xor_bitwise_xor
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

const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=0b1100; Atomics.xor(ia,0,0b1010); __check(__line(ia[0]), "6");
