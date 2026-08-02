// vybe-test: js/tagged_template_deep/tagged_template_with_expression_values
// origin: languages/js/tests/js/test_tagged_template_deep.rs

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

const id = (s, ...v) => String.raw({ raw: s }, ...v);
const a = 3, b = 4;
const result = id`${a} + ${b} = ${a + b}`;
__check(__line(result), "3 + 4 = 7");
