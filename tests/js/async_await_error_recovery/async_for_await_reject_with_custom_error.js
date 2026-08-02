// vybe-test: js/async_await_error_recovery/async_for_await_reject_with_custom_error
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

async function main(){const o=[];class E extends Error{}async function* g(){yield 1;throw new E("custom");}try{for await(const v of g())o.push(v);}catch(e){o.push(e instanceof E?"e:custom":"other");}console.log(o.join(","));}main();
