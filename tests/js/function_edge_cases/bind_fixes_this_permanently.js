// vybe-test: js/function_edge_cases/bind_fixes_this_permanently
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

const obj = { val: 100 };
function getVal() { return this.val; }
const bound = getVal.bind(obj);
__check(__line(bound()), "100");
