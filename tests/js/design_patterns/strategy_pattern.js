// vybe-test: js/design_patterns/strategy_pattern
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

class Sorter {
    constructor(strategy) { this.strategy = strategy; }
    sort(arr) { return this.strategy([...arr]); }
}
const ascending = arr => arr.sort((a, b) => a - b);
const descending = arr => arr.sort((a, b) => b - a);
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
const asc = new Sorter(ascending);
const desc = new Sorter(descending);
__check(__line(asc.sort(nums).join(",")), "1,1,2,3,4,5,6,9");
__check(__line(desc.sort(nums).join(",")), "9,6,5,4,3,2,1,1");
