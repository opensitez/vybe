// vybe-test: js/class_patterns/class_iterable_with_symbol_iterator
// origin: languages/js/tests/js/test_class_patterns.rs

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

class NumberRange {
    constructor(start, end) { this.start = start; this.end = end; }
    [Symbol.iterator]() {
        let current = this.start;
        let end = this.end;
        return {
            next() {
                if (current <= end) return { value: current++, done: false };
                return { done: true };
            }
        };
    }
}
let nums = [...new NumberRange(1, 5)];
__check(__line(nums.join(",")), "1,2,3,4,5");
