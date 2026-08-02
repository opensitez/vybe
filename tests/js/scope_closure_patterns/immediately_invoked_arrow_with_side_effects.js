// vybe-test: js/scope_closure_patterns/immediately_invoked_arrow_with_side_effects
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

const result = (() => {
    const a = 1, b = 2, c = 3;
    return { a, b, c, sum: a + b + c };
})();
__check(__line(result.sum), "6");
__check(__line(result.a), "1");
