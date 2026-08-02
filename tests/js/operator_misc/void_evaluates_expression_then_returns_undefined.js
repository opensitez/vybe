// vybe-test: js/operator_misc/void_evaluates_expression_then_returns_undefined
// origin: languages/js/tests/js/test_operator_misc.rs

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

let seen = [];
const out = void seen.push("side-effect");
__check(__line(String(out)), "undefined");
__check(__line(seen.length), "1");
