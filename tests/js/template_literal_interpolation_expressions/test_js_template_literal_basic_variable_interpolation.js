// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_basic_variable_interpolation
// origin: languages/js/tests/js/test_js_template_literal_interpolation_expressions.rs

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

const name = "Alice";
const age = 30;
__check(__line(`User ${name} is ${age} years old.`), "User Alice is 30 years old.");
