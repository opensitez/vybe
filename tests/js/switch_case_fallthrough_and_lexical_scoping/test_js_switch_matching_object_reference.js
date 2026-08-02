// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_matching_object_reference
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

const obj1 = { id: 1 };
const obj2 = { id: 1 };
let matched = false;
switch(obj1) {
    case obj2: matched = false; break;
    case obj1: matched = true; break;
}
__check(__line(matched), "true");
