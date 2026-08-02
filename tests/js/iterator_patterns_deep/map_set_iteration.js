// vybe-test: js/iterator_patterns_deep/map_set_iteration
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

const map = new Map([["a", 1], ["b", 2]]);
const keyIter = map.keys();
__check(__line(keyIter.next().value), "a");
const set = new Set([10, 20, 30]);
const setIter = set[Symbol.iterator]();
__check(__line(setIter.next().value), "10");
__check(__line([...set.values()].join(",")), "10,20,30");
