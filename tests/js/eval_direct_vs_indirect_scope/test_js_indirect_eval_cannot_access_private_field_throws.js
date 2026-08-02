// vybe-test: js/eval_direct_vs_indirect_scope/test_js_indirect_eval_cannot_access_private_field_throws
// origin: languages/js/tests/js/test_js_eval_direct_vs_indirect_scope.rs

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
    #code = "1234";
    getCode() {
        return (0, eval)("this.#code");
    }
}
try {
    new Secret().getCode();
} catch (e) {
    __check(__line("Indirect Eval Private Field SyntaxError"), "Indirect Eval Private Field SyntaxError");
}
