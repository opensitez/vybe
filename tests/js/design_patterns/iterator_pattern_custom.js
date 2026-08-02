// vybe-test: js/design_patterns/iterator_pattern_custom
// origin: languages/js/tests/js/test_design_patterns.rs

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
                    ? { value: cur, done: false, ...{ next: () => { cur += step; } } }
                    : { done: true };
            }
        };
    }
}
// Use a generator instead for cleaner implementation
function* range(start, end, step = 1) {
    for (let i = start; i < end; i += step) yield i;
}
console.log([...range(0, 10, 2)].join(","));
console.log([...range(1, 6)].join(","));
