// vybe-test: js/typed_array_constructors_matrix/typed_array_iterator_next_returns_value
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

const it=new Uint8Array([4,5]).values(); __check(__line(it.next().value), "4");
