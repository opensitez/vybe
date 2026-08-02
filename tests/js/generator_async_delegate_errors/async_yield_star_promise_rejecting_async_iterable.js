// vybe-test: js/generator_async_delegate_errors/async_yield_star_promise_rejecting_async_iterable
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

async function* bad(){yield 1;await Promise.reject("nope");} async function* wrap(){try{yield* bad();}catch(e){yield "e:"+e;}} (async()=>{const a=[];for await(const v of wrap())a.push(v);console.log(a.join(","));})();
