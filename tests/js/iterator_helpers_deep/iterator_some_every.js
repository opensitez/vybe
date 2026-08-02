// vybe-test: js/iterator_helpers_deep/iterator_some_every
// origin: languages/js/tests/js/test_iterator_helpers_deep.rs

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

function someIter(iter, pred) {
    for (const v of iter) if (pred(v)) return true;
    return false;
}
function everyIter(iter, pred) {
    for (const v of iter) if (!pred(v)) return false;
    return true;
}
const nums = [2, 4, 6, 7, 8];
console.log(someIter(nums, x => x % 2 !== 0));
console.log(everyIter([2, 4, 6], x => x % 2 === 0));
