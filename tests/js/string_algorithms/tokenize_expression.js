// vybe-test: js/string_algorithms/tokenize_expression
// origin: languages/js/tests/js/test_string_algorithms.rs

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

function tokenize(expr) {
    return expr.match(/\d+|[+\-*/()]/g) || [];
}
__check(__line(tokenize("3+4*(2-1)").join(",")), "3,+,4,*,(,2,-,1,)");
__check(__line(tokenize("100/25+5").join(",")), "100,/,25,+,5");
