// vybe-test: js/json_semantics_matrix/json_roundtrip_with_replacer_and_reviver_can_recover_numbers
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

const src = { a: 2, b: 4 };
const json = JSON.stringify(src, (key, value) => {
    return typeof value === "number" ? value * 2 : value;
});
const back = JSON.parse(json, (key, value) => {
    return typeof value === "number" ? value / 2 : value;
});
__check(__line(back.a), "2");
__check(__line(back.b), "4");
