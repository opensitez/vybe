// vybe-test: js/map_set/closure_mutate_on_object
// origin: languages/js/tests/js/js_map_set_test.rs

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

function make() {
            let n = 0;
            return {
                inc: () => { n = n + 1; return n; },
                getN: () => n
            };
        }
        let c = make();
        c.inc();
        c.inc();
        c.inc();
        __check(__line(c.getN()), "3");
