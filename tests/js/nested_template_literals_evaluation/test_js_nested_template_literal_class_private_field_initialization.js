// vybe-test: js/nested_template_literals_evaluation/test_js_nested_template_literal_class_private_field_initialization
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

class Secret {
    #code = `Priv_${`Secret_${99}`}`;
    getCode() { return this.#code; }
}
__check(__line(new Secret().getCode()), "Priv_Secret_99");
