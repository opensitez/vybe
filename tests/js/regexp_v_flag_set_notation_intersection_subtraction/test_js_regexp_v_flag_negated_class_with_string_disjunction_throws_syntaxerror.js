// vybe-test: js/regexp_v_flag_set_notation_intersection_subtraction/test_js_regexp_v_flag_negated_class_with_string_disjunction_throws_syntaxerror
// origin: languages/js/tests/js/test_js_regexp_v_flag_set_notation_intersection_subtraction.rs

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

try {
    eval("const re = /[^\\q{abc}]/v;"); // Negated character classes cannot contain string sequences!
} catch (e) {
    __check(__line("Negated Sequence Class SyntaxError"), "Negated Sequence Class SyntaxError");
}
