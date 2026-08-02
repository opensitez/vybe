// vybe-test: js/template_literal_advanced/template_literals_escape_dollar_brace
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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

const marker = "${";
__check(__line(`literal ${marker}value`), "literal ${value");
__check(__line(`line break escapes: \\n`), "line break escapes: \\n");
