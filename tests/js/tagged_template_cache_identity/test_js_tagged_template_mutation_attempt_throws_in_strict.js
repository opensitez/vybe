// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_mutation_attempt_throws_in_strict
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

function tag(strings) {
    "use strict";
    try {
        strings[0] = "Mutated";
    } catch (e) {
        __check(__line("Mutation TypeError"), "Mutation TypeError");
    }
}
tag`Original`;
