// vybe-test: js/template_literal_advanced/tagged_template_transform_values
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

function upper(strings, ...values) {
    return strings.reduce((acc, str, i) => {
        const val = values[i-1];
        return acc + (val !== undefined ? String(val).toUpperCase() : "") + str;
    });
}
const name = "alice";
__check(__line(upper`Hello, ${name}!`), "Hello, ALICE!");
