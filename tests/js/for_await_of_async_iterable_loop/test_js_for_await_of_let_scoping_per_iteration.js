// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_let_scoping_per_iteration
// origin: languages/js/tests/js/test_js_for_await_of_async_iterable_loop.rs

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

(async () => {
    const fns = [];
    for await (const i of [1, 2, 3]) {
        fns.push(() => i);
    }
    console.log(fns.map(f => f()).join(","));
})();
