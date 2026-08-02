// vybe-test: js/typed_array_constructors_matrix/typed_array_entries_yields_index_value_pairs
// origin: languages/js/tests/js/test_typed_array_constructors_matrix.rs

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

const e=new Uint8Array([7]).entries().next().value; __check(__line(e[0]), "0");__check(__line(e[1]), "7");
