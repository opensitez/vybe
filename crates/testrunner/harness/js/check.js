// Vybe test harness — JavaScript.
//
// The JS counterpart of harness/go/check.go, and the direct analogue of
// test262's harness/assert.js: ordinary source in the language under test, so
// it can be linted, formatted and debugged with that language's own tools.
//
// A test's verdict is its EXIT CODE. `__check` prints its own diagnostic
// BEFORE throwing, because an uncaught error surfaces as `RuntimeError:
// [object]` — 1,692 of testecma's 2,158 failures say exactly that and nothing
// more. The printed line survives that gap on every runtime.

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
