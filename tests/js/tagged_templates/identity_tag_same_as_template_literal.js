// vybe-test: js/tagged_templates/identity_tag_same_as_template_literal
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

function id(strings, ...values) {
    return String.raw({ raw: strings }, ...values);
}
const x = 5;
__check(__line(id`value is ${x}`), "value is 5");
