// vybe-test: js/function_edge_cases/function_expression_not_hoisted
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

let result;
try {
    result = notHoisted();
} catch (e) {
    result = "error:" + e.constructor.name;
}
var notHoisted = function() { return "late"; };
__check(__line(result), "error:TypeError");
