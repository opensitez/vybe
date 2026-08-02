// vybe-test: js/json_semantics_matrix/json_stringify_sparse_array_holes_become_null
// origin: languages/js/tests/js/test_json_semantics_matrix.rs

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

const arr = [];
arr[1] = 2;
arr[3] = 4;
__check(__line(JSON.stringify(arr)), "[null,2,null,4]");
