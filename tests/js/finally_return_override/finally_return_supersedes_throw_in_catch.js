// vybe-test: js/finally_return_override/finally_return_supersedes_throw_in_catch
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

function f(){try{throw 1;}catch{throw 2;}finally{return 9;}}__check(__line(f()), "9");
