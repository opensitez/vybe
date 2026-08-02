// vybe-test: js/json_semantics_matrix/json_stringify_escapes_newline_and_tab
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

const json = JSON.stringify({ s: "a\n\tb" });
__check(__line(json.indexOf("\\n") >= 0), "true");
__check(__line(json.indexOf("\\t") >= 0), "true");
