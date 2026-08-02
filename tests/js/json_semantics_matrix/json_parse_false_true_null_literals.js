// vybe-test: js/json_semantics_matrix/json_parse_false_true_null_literals
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

const obj = JSON.parse('{"t":true,"f":false,"n":null}');
__check(__line(obj.t), "true");
__check(__line(obj.f), "false");
__check(__line(obj.n === null), "true");
