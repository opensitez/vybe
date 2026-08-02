// vybe-test: js/nested_template_literals_evaluation/test_js_nested_template_literal_expression_evaluation_order
// origin: languages/js/tests/js/test_js_nested_template_literals_evaluation.rs

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

let order = [];
function first() { order.push(1); return "1"; }
function second() { order.push(2); return "2"; }

const str = `First: ${first()} (${`Second: ${second()}`})`;
__check(__line(str + "|Order=" + order.join(",")), "First: 1 (Second: 2)|Order=1,2");
