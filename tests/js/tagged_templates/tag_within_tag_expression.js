// vybe-test: js/tagged_templates/tag_within_tag_expression
// origin: languages/js/tests/js/test_tagged_templates.rs

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

function double(strings, ...values) {
    return values[0] * 2;
}
function outer(strings, ...values) {
    return "result:" + values[0];
}
__check(__line(outer`val=${double`${21}`}`), "result:42");
