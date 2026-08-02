// vybe-test: js/interop/test_a08_arrow_in_method_captures_this
// origin: languages/js/tests/js/js_interop_test.rs

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
            constructor() { this.ticks = 0; }
            start() {
                let tick = () => { this.ticks++; };
                tick();
                tick();
                tick();
                return this.ticks;
            }
        }
        let t = new Timer();
        __check(__line(t.start()), "3");
