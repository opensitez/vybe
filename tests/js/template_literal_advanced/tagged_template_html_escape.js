// vybe-test: js/template_literal_advanced/tagged_template_html_escape
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

function html(strings, ...values) {
    function escape(s) {
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }
    return strings.reduce((acc, str, i) =>
        acc + (i > 0 ? escape(values[i-1]) : "") + str
    );
}
const user = "<script>alert(1)</script>";
__check(__line(html`Hello ${user}!`), "Hello &lt;script&gt;alert(1)&lt;/script&gt;!");
