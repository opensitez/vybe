// vybe-test: js/atomics_operations_matrix/atomics_operations_on_uint32_array
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

const sab=new SharedArrayBuffer(4); const ia=new Uint32Array(sab); Atomics.store(ia,0,100); __check(__line(Atomics.add(ia,0,50)), "100");__check(__line(ia[0]), "150");
