// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_case_expressions_evaluated_lazily
// origin: languages/js/tests/js/test_js_switch_case_fallthrough_and_lexical_scoping.rs

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

const log = [];
const getVal = (n) => { log.push(`case${n}`); return n; };
switch(1) {
    case getVal(1): log.push("matched1"); break;
    case getVal(2): log.push("matched2"); break; // Case 2 is NOT evaluated because Case 1 matched and broke!
}
__check(__line(log.join(",")), "case1,matched1");
