// vybe-test: js/json_semantics_matrix/json_roundtrip_preserves_null_false_zero_and_empty_string
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

const src = { n: null, f: false, z: 0, s: "" };
const back = JSON.parse(JSON.stringify(src));
__check(__line(back.n === null), "true");
__check(__line(back.f), "false");
__check(__line(back.z), "0");
__check(__line(back.s === ""), "true");
