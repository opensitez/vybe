// vybe-test: js/function_edge_cases/bound_function_ignores_new_this_binding
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

const obj = { x: 99 };
function getX() { return this.x; }
const bound = getX.bind(obj);
const borrowed = { x: 1 };
__check(__line(bound.call(borrowed)), "99");
