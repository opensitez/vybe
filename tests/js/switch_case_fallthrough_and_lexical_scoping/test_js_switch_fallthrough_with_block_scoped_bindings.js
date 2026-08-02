// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_fallthrough_with_block_scoped_bindings
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
switch("x") {
    case "x": {
        const marker = "A";
        out.push(marker);
    }
    case "y": {
        const marker = "B";
        out.push(marker);
        break;
    }
}
__check(__line(out.join("|")), "A|B");
