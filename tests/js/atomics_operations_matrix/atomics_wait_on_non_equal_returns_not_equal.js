// vybe-test: js/atomics_operations_matrix/atomics_wait_on_non_equal_returns_not_equal
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

const sab=new SharedArrayBuffer(4); const ia=new Int32Array(sab); ia[0]=1; __check(__line(Atomics.wait(ia,0,2,0)), "not-equal");
