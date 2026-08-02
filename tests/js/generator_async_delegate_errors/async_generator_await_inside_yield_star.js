// vybe-test: js/generator_async_delegate_errors/async_generator_await_inside_yield_star
// origin: languages/js/tests/js/test_generator_async_delegate_errors.rs

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

async function* inner(){yield await Promise.resolve(2);} async function* outer(){yield* inner();} (async()=>{const v=await outer().next();console.log(v.value);})();
