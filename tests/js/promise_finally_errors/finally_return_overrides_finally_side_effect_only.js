// vybe-test: js/promise_finally_errors/finally_return_overrides_finally_side_effect_only
// origin: languages/js/tests/js/test_promise_finally_errors.rs

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

const o=[];Promise.resolve(1).finally(()=>{o.push("s");return "r";}).then(v=>console.log(v+o.join("")));
