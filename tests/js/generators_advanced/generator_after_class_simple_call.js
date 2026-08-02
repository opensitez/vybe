// vybe-test: js/generators_advanced/generator_after_class_simple_call
// origin: languages/js/tests/js/test_generators_advanced.rs

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

class Range {
    constructor(start, end, step = 1) {
        this.start = start; this.end = end; this.step = step;
    }
    [Symbol.iterator]() {
        let cur = this.start;
        const { end, step } = this;
        return {
            next() {
                return cur < end
                    ? { value: cur, done: false }
                    : { done: true };
            }
        };
    }
}
function* range(start, end, step = 1) {
    for (let i = start; i < end; i += step) yield i;
}
console.log([...range(0, 5, 1)].join(","));
