// vybe-test: js/closure_scope_deep_patterns/nested_function_closure_shared_state
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

function makeShared() {
    let shared = [];
    function add(v) { shared.push(v); }
    function get() { return [...shared]; }
    function clear() { shared = []; }
    return { add, get, clear };
}
const { add, get, clear } = makeShared();
add(1); add(2); add(3);
__check(__line(get().join(",")), "1,2,3");
clear();
__check(__line(get().length), "0");
