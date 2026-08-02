// vybe-test: js/typed_array_constructors_matrix/float32array_map_to_plain_array
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

const a=new Float32Array([2,3]); const m=Array.from(a,x=>x*2); __check(__line(m.join(",")), "4,6");
