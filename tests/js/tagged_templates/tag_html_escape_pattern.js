// vybe-test: js/tagged_templates/tag_html_escape_pattern
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

function html(strings, ...values) {
    const escape = v => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    return strings.reduce((acc, s, i) => acc + s + (values[i] !== undefined ? escape(values[i]) : ""), "");
}
const user = "<script>alert('xss')</script>";
__check(__line(html`<p>${user}</p>`), "<p>&lt;script&gt;alert('xss')&lt;/script&gt;</p>");
