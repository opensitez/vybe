// vybe-test: js/template_literal_advanced/template_as_heredoc_style
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

function dedent(str) {
    const lines = str.split("\n").filter(l => l.trim());
    const indent = Math.min(...lines.map(l => l.match(/^\s*/)[0].length));
    return lines.map(l => l.slice(indent)).join("\n");
}
const code = dedent(`
    function hello() {
        return "world";
    }
`);
__check(__line(code.startsWith("function")), "true");
