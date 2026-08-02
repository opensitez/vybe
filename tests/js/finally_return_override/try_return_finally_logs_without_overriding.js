// vybe-test: js/finally_return_override/try_return_finally_logs_without_overriding
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

function f(){const o=[];try{return "ok";}finally{o.push("f");}return o.join(",");}__check(__line(f()), "ok");
