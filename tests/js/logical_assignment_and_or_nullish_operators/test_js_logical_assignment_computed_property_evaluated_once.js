// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_computed_property_evaluated_once
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

let keyEvaluationCount = 0;
const getKey = () => { keyEvaluationCount++; return "prop"; };
const obj = { prop: null };

obj[getKey()] ??= "Assigned";
__check(__line(obj.prop + "|KeyEvalCount=" + keyEvaluationCount), "Assigned|KeyEvalCount=1");
