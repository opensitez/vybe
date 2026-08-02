// vybe-test: js/ecma_async/async_method_can_use_this_before_await
// origin: languages/js/tests/js/test_ecma_async.rs

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

class Counter {
    constructor() { this.value = 2; }
    async double() {
        console.log(this.value);
        return this.value * 2;
    }
}
const c = new Counter();
console.log(c.double());
