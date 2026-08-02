// vybe-test: js/string_processing_deep/template_engine_simple
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function render(template, data) {
    return template.replace(/\{\{(\w+)\}\}/g, (_, key) => data[key] ?? "");
}
const tmpl = "Hello, {{name}}! You are {{age}} years old.";
__check(__line(render(tmpl, { name: "Alice", age: 30 })), "Hello, Alice! You are 30 years old.");
__check(__line(render("{{missing}} world", {})), " world");
