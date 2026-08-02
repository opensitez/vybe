// vybe-test: js/finally_return_override/nested_finally_inner_return_wins_over_outer
// origin: languages/js/tests/js/test_finally_return_override.rs

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

function f(){try{try{return "inner";}finally{return "fin";}}finally{return "outer";}}__check(__line(f()), "outer");
