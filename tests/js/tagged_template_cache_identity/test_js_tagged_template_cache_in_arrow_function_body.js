// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_in_arrow_function_body
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

const tag = strings => strings;
const getFn = () => tag`ArrowTemplate`;
__check(__line(getFn() === getFn()), "true");
