// vybe-test: js/string_processing_patterns/string_interpolation_template
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

function interpolate(template, data) {
    return template.replace(/\{\{(\w+)\}\}/g, (_, key) => data[key] ?? "");
}
const result = interpolate("Hello {{name}}, you have {{count}} messages!", {
    name: "Alice",
    count: 5
});
__check(__line(result), "Hello Alice, you have 5 messages!");
