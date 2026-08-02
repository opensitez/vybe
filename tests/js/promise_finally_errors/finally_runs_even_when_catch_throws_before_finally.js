// vybe-test: js/promise_finally_errors/finally_runs_even_when_catch_throws_before_finally
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

const o=[];Promise.reject("a").catch(()=>{throw "c";}).finally(()=>o.push("f")).catch(e=>o.push("e:"+e)).then(()=>console.log(o.join(",")));
