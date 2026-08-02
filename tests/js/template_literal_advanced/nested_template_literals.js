// vybe-test: js/template_literal_advanced/nested_template_literals
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

const items = ["a", "b", "c"];
const result = `list: ${items.map(i => `[${i}]`).join(",")}`;
__check(__line(result), "list: [a],[b],[c]");
