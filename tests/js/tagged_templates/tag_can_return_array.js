// vybe-test: js/tagged_templates/tag_can_return_array
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

function tokens(strings, ...values) {
    const result = [];
    strings.forEach((s, i) => {
        if (s) result.push(s);
        if (i < values.length) result.push(values[i]);
    });
    return result;
}
const a = 1, b = 2;
const parts = tokens`A${a}B${b}C`;
console.log(parts.join("-"));
