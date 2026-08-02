// vybe-test: js/scope_closure_patterns/nested_closure_scope_chain
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

const a = 1;
function level1() {
    const b = 2;
    function level2() {
        const c = 3;
        function level3() {
            return a + b + c;
        }
        return level3();
    }
    return level2();
}
__check(__line(level1()), "6");
