// vybe-test: js/function_edge_cases/function_declaration_shadows_global
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

function f() {
    function Math() { return "custom"; }
    return Math();
}
__check(__line(f()), "custom");
