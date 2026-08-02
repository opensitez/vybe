// vybe-test: js/generator_async_delegate_errors/async_for_await_yield_star_async_gen
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

async function* nums(){yield 3;yield 4;} async function* wrap(){yield* nums();} (async()=>{let s=0;for await(const v of wrap())s+=v;console.log(s);})();
