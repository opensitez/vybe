// vybe-test: js/iterator_patterns_deep/custom_iterable_class
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

class Matrix {
    constructor(rows) { this.rows = rows; }
    [Symbol.iterator]() {
        let r = 0, c = 0;
        const rows = this.rows;
        return {
            next() {
                if (r >= rows.length) return { done: true };
                const value = rows[r][c++];
                if (c >= rows[r].length) { c = 0; r++; }
                return { value, done: false };
            }
        };
    }
}
const m = new Matrix([[1,2],[3,4],[5,6]]);
__check(__line([...m].join(",")), "1,2,3,4,5,6");
