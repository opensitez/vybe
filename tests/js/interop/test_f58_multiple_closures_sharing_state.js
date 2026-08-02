// vybe-test: js/interop/test_f58_multiple_closures_sharing_state
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

function makeStore() {
            let value = 0;
            return {
                set(v) { value = v; },
                get() { return value; },
                inc() { value++; }
            };
        }
        let store = makeStore();
        store.set(10);
        store.inc();
        store.inc();
        __check(__line(store.get()), "12");
