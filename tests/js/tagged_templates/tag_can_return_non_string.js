// vybe-test: js/tagged_templates/tag_can_return_non_string
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

function tag(strings, ...values) {
    return values.reduce((a, b) => a + b, 0);
}
const result = tag`x${10}y${20}z${30}`;
__check(__line(result), "60");
