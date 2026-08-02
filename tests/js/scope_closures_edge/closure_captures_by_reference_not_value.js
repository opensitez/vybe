// vybe-test: js/scope_closures_edge/closure_captures_by_reference_not_value
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

let x = 1;
const get = () => x;
const set = v => { x = v; };
__check(__line(get()), "1");
set(42);
__check(__line(get()), "42");
