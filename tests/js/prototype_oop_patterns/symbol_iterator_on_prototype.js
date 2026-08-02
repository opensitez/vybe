// vybe-test: js/prototype_oop_patterns/symbol_iterator_on_prototype
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

function Range(start, end) { this.start = start; this.end = end; }
Range.prototype[Symbol.iterator] = function*() {
    for (let i = this.start; i <= this.end; i++) yield i;
};
const r = new Range(1, 5);
console.log([...r].join(","));
console.log(Array.from(r).join(","));
