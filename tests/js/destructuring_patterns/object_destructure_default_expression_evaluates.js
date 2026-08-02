// vybe-test: js/destructuring_patterns/object_destructure_default_expression_evaluates
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

let calls = 0;
function def() { calls++; return 42; }
const { x = def(), y = def() } = { x: 1 };
__check(__line(x), "1");   // 1 — def() not called
__check(__line(y), "42");   // 42 — def() called
__check(__line(calls), "1"); // 1
