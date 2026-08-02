// vybe-test: js/iterator_patterns_deep/manual_iterator_usage
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

const arr = [10, 20, 30];
const iter = arr[Symbol.iterator]();
__check(__line(iter.next().value), "10");
__check(__line(iter.next().value), "20");
__check(__line(iter.next().done), "false");
__check(__line(iter.next().done), "true");
