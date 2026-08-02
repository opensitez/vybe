// vybe-test: js/ecma_functions/arrow_function_lexical_this
// origin: languages/js/tests/js/test_ecma_functions.rs

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

const counter = {
    count: 0,
    inc() {
        const step = () => {
            this.count += 1;
        };
        step();
        step();
        return this.count;
    }
};
__check(__line(counter.inc()), "2");
__check(__line(counter.count), "2");
