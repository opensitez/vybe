// vybe-test: js/generator_async_delegate_errors/async_generator_return_stops_iteration
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

async function* g(){yield 1;return "end";yield 2;} (async()=>{const a=[];for await(const v of g())a.push(v);console.log(a.length);})();
