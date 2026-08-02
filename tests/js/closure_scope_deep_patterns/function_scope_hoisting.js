// vybe-test: js/closure_scope_deep_patterns/function_scope_hoisting
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

__check(__line(hoisted()), "hoisted");  // fn declarations hoisted
function hoisted() { return "hoisted"; }
var x = 10;
function useX() { return x; }
__check(__line(useX()), "10");
