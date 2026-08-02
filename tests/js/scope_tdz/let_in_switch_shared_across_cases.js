// vybe-test: js/scope_tdz/let_in_switch_shared_across_cases
// origin: languages/js/tests/js/test_scope_tdz.rs

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

let result = "";
switch (1) {
    case 1:
        let x = "shared";
        result += x;
    case 2:
        result += "-done";
}
__check(__line(result), "shared-done");
