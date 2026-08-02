// vybe-test: js/iterator_protocol/object_is_self_iterating_via_symbol_iterator
// origin: languages/js/tests/js/test_iterator_protocol.rs

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
    constructor(start, end) { this.start = start; this.end = end; }
    [Symbol.iterator]() {
        let cur = this.start;
        const end = this.end;
        return {
            next() {
                return cur <= end
                    ? { value: cur++, done: false }
                    : { done: true };
            }
        };
    }
}
__check(__line([...new Range(3, 6)].join(",")), "3,4,5,6");
