// vybe-test: js/atomics_operations_matrix/atomics_notify_no_waiters_returns_zero
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

const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); __check(__line(Atomics.notify(ia,0,1)), "0");
