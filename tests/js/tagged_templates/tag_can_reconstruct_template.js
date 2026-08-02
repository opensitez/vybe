// vybe-test: js/tagged_templates/tag_can_reconstruct_template
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

function tag(strings, ...values) {
    return strings.reduce((acc, s, i) => acc + s + (values[i] !== undefined ? values[i] : ""), "");
}
const a = "hello";
const b = "world";
__check(__line(tag`${a}, ${b}!`), "hello, world!");
