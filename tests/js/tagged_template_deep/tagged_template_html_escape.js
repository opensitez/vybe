// vybe-test: js/tagged_template_deep/tagged_template_html_escape
// origin: languages/js/tests/js/test_tagged_template_deep.rs

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
        const val = values[i - 1] != null
            ? String(values[i - 1]).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
            : "";
        return acc + val + str;
    });
}
const user = "<script>alert(1)</script>";
const result = html`<p>Hello ${user}!</p>`;
__check(__line(result.includes("&lt;script&gt;")), "true");
