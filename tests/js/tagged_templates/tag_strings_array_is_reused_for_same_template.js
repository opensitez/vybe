// vybe-test: js/tagged_templates/tag_strings_array_is_reused_for_same_template
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

let firstRef;
function tag(strings) {
    if (!firstRef) firstRef = strings;
    return strings === firstRef;
}
function call() { return tag`hello`; }
const r1 = call();
const r2 = call();
__check(__line(r1), "true");
__check(__line(r2), "true");
