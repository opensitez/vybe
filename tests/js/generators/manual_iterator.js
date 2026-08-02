// vybe-test: js/generators/manual_iterator
// origin: languages/js/tests/js/test_generators.rs

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

let iter = {
    items: ["x", "y", "z"],
    index: 0,
    next() {
        if (this.index < this.items.length) {
            return { value: this.items[this.index++], done: false };
        }
        return { done: true };
    },
    [Symbol.iterator]() { return this; }
};
for (let v of iter) console.log(v);
