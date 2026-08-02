// vybe-test: js/async_await_error_recovery/async_outer_catch_on_inner_reraise
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

async function main(){try{try{await Promise.reject("a");}catch(e){throw "b:"+e;}}catch(e){console.log(e);}}main();
