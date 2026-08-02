// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_default_case_placement_in_middle_with_fallthrough
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

const log = [];
switch(99) {
    case 1: log.push("c1"); break;
    default: log.push("def"); // Default matches, falls through to case 2!
    case 2: log.push("c2"); break;
}
__check(__line(log.join(",")), "def,c2");
