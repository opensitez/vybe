// vybe-test: js/async_await_error_recovery/async_catch_with_conditional_reraise
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

async function main(){try{await Promise.reject("code:404");}catch(e){if(String(e).includes("404"))console.log("handled");else throw e;}}main();
