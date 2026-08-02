// vybe-test: js/json_semantics_matrix/json_stringify_then_parse_preserves_property_order
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

const src = {};
src.c = 1;
src.a = 2;
const back = JSON.parse(JSON.stringify(src));
__check(__line(Object.keys(back).join(",")), "c,a");
