// vybe-test: js/async_await_error_recovery/async_for_await_with_continue
// origin: languages/js/tests/js/test_async_await_error_recovery.rs

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

async function main(){const o=[];async function* g(){yield 1;yield 2;yield 3;}for await(const v of g()){if(v===2)continue;o.push(v);}console.log(o.join(","));}main();
