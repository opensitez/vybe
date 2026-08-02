// vybe-test: js/with_statement_unscopables_protocol/with_target_expression_evaluated_once
// origin: languages/js/tests/js/test_js_with_statement_unscopables_protocol.rs

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

let count = 0;
function getObj() { count++; return { a: 1 }; }
with (getObj()) {
    __check(__line(a + count), "2");
}
