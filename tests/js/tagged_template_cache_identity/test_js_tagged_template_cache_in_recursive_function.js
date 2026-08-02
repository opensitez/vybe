// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_in_recursive_function
// origin: languages/js/tests/js/test_js_tagged_template_cache_identity.rs

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

function tag(strings) { return strings; }
function recurse(n) {
    const template = tag`RecurseNode`;
    if (n <= 0) return [template];
    return [template, ...recurse(n - 1)];
}
const list = recurse(2);
__check(__line(list[0] === list[1] && list[1] === list[2]), "true");
