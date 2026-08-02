// vybe-test: js/functional_fp_patterns/transducer_pattern
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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

const map = fn => reducer => (acc, val) => reducer(acc, fn(val));
const filter = pred => reducer => (acc, val) => pred(val) ? reducer(acc, val) : acc;
const append = (acc, val) => { acc.push(val); return acc; };

const xform = [
    filter(x => x % 2 === 0),
    map(x => x * x)
].reduce((a, b) => b(a), append);

const result = [1,2,3,4,5,6].reduce(xform, []);
__check(__line(result.join(",")), "4,16,36");
