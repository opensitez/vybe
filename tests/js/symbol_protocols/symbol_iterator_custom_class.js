// vybe-test: js/symbol_protocols/symbol_iterator_custom_class
// origin: languages/js/tests/js/test_symbol_protocols.rs

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
    constructor(start, end, step = 1) {
        this.start = start; this.end = end; this.step = step;
    }
    [Symbol.iterator]() {
        let current = this.start;
        const { end, step } = this;
        return {
            next() {
                if (current <= end) { const value = current; current += step; return { value, done: false }; }
                return { value: undefined, done: true };
            }
        };
    }
}
const r = new NumberRange(1, 10, 2);
__check(__line([...r].join(",")), "1,3,5,7,9");
const arr2 = [...new NumberRange(10, 50, 10)];
__check(__line(arr2[0]), "10");
__check(__line(arr2[2]), "30");
