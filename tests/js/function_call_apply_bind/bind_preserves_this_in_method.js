// vybe-test: js/function_call_apply_bind/bind_preserves_this_in_method
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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

class Timer {
    constructor() { this.count = 0; }
    tick() { this.count++; return this.count; }
}
const t = new Timer();
const tick = t.tick.bind(t);
tick();
tick();
const result = tick();
__check(__line(result), "3");
__check(__line(t.count), "3");
