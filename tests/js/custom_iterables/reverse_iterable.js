// vybe-test: js/custom_iterables/reverse_iterable
// origin: languages/js/tests/js/test_custom_iterables.rs

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

class ReverseIterable {
    constructor(arr) { this.arr = arr; }
    [Symbol.iterator]() {
        const arr = this.arr;
        let i = arr.length - 1;
        return {
            next() {
                return i >= 0 ? { value: arr[i--], done: false } : { done: true };
            }
        };
    }
}
const rev = new ReverseIterable([1, 2, 3, 4, 5]);
__check(__line([...rev].join(",")), "5,4,3,2,1");
