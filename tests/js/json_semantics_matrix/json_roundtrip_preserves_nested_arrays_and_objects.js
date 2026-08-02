// vybe-test: js/json_semantics_matrix/json_roundtrip_preserves_nested_arrays_and_objects
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

const src = { a: [{ b: 2 }, 3] };
const back = JSON.parse(JSON.stringify(src));
__check(__line(back.a[0].b), "2");
__check(__line(back.a[1]), "3");
