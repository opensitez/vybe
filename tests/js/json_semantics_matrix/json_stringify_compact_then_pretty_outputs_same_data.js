// vybe-test: js/json_semantics_matrix/json_stringify_compact_then_pretty_outputs_same_data
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

const src = { a: 1, b: [2, 3] };
const a = JSON.parse(JSON.stringify(src));
const b = JSON.parse(JSON.stringify(src, null, 2));
__check(__line(a.b[1] === b.b[1]), "true");
__check(__line(JSON.stringify(a) === JSON.stringify(b)), "true");
