// vybe-test: js/iterator_helpers_es2025/iterator_map_is_lazy
// origin: languages/js/tests/js/test_iterator_helpers_es2025.rs

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

const called = [];
const iter = Iterator.from([1, 2, 3]).map(x => { called.push(x); return x * 2; });
__check(__line(called.length), "0");
iter.next();
__check(__line(called.length), "1");
