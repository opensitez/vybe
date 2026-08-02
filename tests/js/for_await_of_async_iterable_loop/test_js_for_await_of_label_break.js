// vybe-test: js/for_await_of_async_iterable_loop/test_js_for_await_of_label_break
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
    const log = [];
    outer: for await (const i of [1, 2]) {
        for await (const j of [10, 20]) {
            if (i === 1 && j === 20) break outer;
            log.push(`${i}:${j}`);
        }
    }
    console.log(log.join(","));
})();
