// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_continue_statement_inside_loop
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
for (let i = 1; i <= 3; i++) {
    switch(i) {
        case 2: continue; // 'continue' inside switch targets enclosing loop!
    }
    log.push(i);
}
console.log(log.join(","));
