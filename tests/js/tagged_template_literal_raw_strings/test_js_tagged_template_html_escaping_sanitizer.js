// vybe-test: js/tagged_template_literal_raw_strings/test_js_tagged_template_html_escaping_sanitizer
// origin: languages/js/tests/js/test_js_tagged_template_literal_raw_strings.rs

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
    return strings.reduce((acc, str, i) => {
        let val = values[i - 1];
        if (typeof val === "string") {
            val = val.replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }
        return acc + val + str;
    });
}
const user = "<script>alert(1)</script>";
__check(__line(html`<div>${user}</div>`), "<div>&lt;script&gt;alert(1)&lt;/script&gt;</div>");
