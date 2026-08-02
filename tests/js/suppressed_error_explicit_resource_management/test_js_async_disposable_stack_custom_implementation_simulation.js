// vybe-test: js/suppressed_error_explicit_resource_management/test_js_async_disposable_stack_custom_implementation_simulation
// origin: languages/js/tests/js/test_js_suppressed_error_explicit_resource_management.rs

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

class AsyncDisposableStack {
    #resources = [];
    use(res) { this.#resources.push(res); return res; }
    async disposeAsync() {
        while (this.#resources.length > 0) {
            const res = this.#resources.pop();
            if (res && typeof res[Symbol.asyncDispose] === "function") {
                await res[Symbol.asyncDispose]();
            }
        }
    }
}
const log = [];
const stack = new AsyncDisposableStack();
stack.use({ async [Symbol.asyncDispose]() { log.push("AR1"); } });
(async () => {
    await stack.disposeAsync();
    console.log(log.join(","));
})();
