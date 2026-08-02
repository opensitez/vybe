// vybe-test: js/wasi/test_timed_operation
// origin: languages/js/tests/js/js_wasi_test.rs

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

let start = clock.now();
        let sum = 0;
        for (let i = 0; i < 10000; i++) { sum = sum + i; }
        let elapsed = clock.now() - start;
        console.log(sum, elapsed >= 0);
