// vybe-test: js/typed_array_constructors_matrix/int16array_slice_copies_elements
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

const a=new Int16Array([1,2,3]); const s=a.slice(1); s[0]=99; __check(__line(a[1]), "2");
