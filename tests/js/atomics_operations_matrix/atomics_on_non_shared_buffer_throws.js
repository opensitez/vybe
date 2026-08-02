// vybe-test: js/atomics_operations_matrix/atomics_on_non_shared_buffer_throws
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

const ia=new Int32Array(1); try{Atomics.wait(ia,0,0,0);}catch(e){__check(__line(e instanceof TypeError), "true");}
