// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_case_fallthrough_without_break
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
switch(1) {
    case 1: log.push("c1");
    case 2: log.push("c2");
    case 3: log.push("c3"); break;
    case 4: log.push("c4");
}
__check(__line(log.join(",")), "c1,c2,c3");
