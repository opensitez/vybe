// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_per_eval_in_loop
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
const templates = [];
for (let i = 0; i < 3; i++) {
    templates.push(tag`LoopTemplate`);
}
console.log(templates[0] === templates[1] && templates[1] === templates[2]);
