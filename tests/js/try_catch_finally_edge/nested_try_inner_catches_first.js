// vybe-test: js/try_catch_finally_edge/nested_try_inner_catches_first
// origin: languages/js/tests/js/test_try_catch_finally_edge.rs

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

let order = [];
try {
    try {
        throw new Error("boom");
    } catch (e) {
        order.push("inner catch");
        throw e; // re-throw
    }
} catch (e) {
    order.push("outer catch");
}
__check(__line(order.join(",")), "inner catch,outer catch");
