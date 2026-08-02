// vybe-test: js/async_await_error_recovery/async_nested_for_await_inner_throw
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

async function main(){const o=[];async function* outer(){yield 1;yield inner();async function* inner(){throw "inner";}}try{for await(const v of outer())o.push(String(v));}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();
