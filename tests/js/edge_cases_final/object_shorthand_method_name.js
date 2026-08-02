// vybe-test: js/edge_cases_final/object_shorthand_method_name
// origin: languages/js/tests/js/test_edge_cases_final.rs

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

const name = "greet";
const obj = {
    [name]() { return "hello"; },
    get value() { return 42; },
};
__check(__line(obj.greet()), "hello");
__check(__line(obj.value), "42");
__check(__line(typeof obj.greet), "function");
