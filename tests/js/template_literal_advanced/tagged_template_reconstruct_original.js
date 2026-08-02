// vybe-test: js/template_literal_advanced/tagged_template_reconstruct_original
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

function identity(strings, ...values) {
    return strings.reduce((acc, str, i) => acc + (values[i-1] ?? "") + str);
}
const x = 42;
__check(__line(identity`value is ${x} done`), "value is 42 done");
