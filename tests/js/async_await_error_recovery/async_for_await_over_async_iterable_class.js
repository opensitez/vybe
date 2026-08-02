// vybe-test: js/async_await_error_recovery/async_for_await_over_async_iterable_class
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

async function main(){const o=[];class It{async*[Symbol.asyncIterator](){yield 1;throw "cls";}}try{for await(const v of new It())o.push(v);}catch(e){o.push("e:"+e);}console.log(o.join(","));}main();
