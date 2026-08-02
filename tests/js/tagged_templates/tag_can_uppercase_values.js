// vybe-test: js/tagged_templates/tag_can_uppercase_values
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

function upper(strings, ...values) {
    return strings.reduce((acc, s, i) => {
        const v = values[i] !== undefined ? String(values[i]).toUpperCase() : "";
        return acc + s + v;
    }, "");
}
const name = "world";
__check(__line(upper`hello ${name}!`), "hello WORLD!");
