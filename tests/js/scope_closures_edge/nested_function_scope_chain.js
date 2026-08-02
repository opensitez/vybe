// vybe-test: js/scope_closures_edge/nested_function_scope_chain
// origin: languages/js/tests/js/test_scope_closures_edge.rs

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

const outer = 1;
function level1() {
    const mid = 2;
    function level2() {
        const inner = 3;
        return outer + mid + inner;
    }
    return level2();
}
__check(__line(level1()), "6");
