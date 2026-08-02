// vybe-test: js/generator_async_delegate_errors/async_generator_return_method_resolves_value
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

(async () => { async function* g() { yield 1; yield 2; } const gen = g(); await gen.next(); const r = await gen.return("custom_ret"); console.log(r.value + "|" + r.done); })();
