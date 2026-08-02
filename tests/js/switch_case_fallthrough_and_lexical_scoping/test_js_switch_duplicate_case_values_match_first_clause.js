// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_duplicate_case_values_match_first_clause
// origin: languages/js/tests/js/test_js_switch_case_fallthrough_and_lexical_scoping.rs

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

const out = [];
switch (3) {
    case 1: out.push("first");
    case 3: out.push("firstMatch"); break;
    case 3: out.push("secondMatch"); break;
    default: out.push("default");
}
__check(__line(out.join("|")), "firstMatch");
