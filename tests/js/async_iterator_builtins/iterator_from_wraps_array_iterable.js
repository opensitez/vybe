// vybe-test: js/async_iterator_builtins/iterator_from_wraps_array_iterable
// origin: languages/js/tests/js/test_async_iterator_builtins.rs

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

const iterator = Iterator.from([10, 20, 30]);
__check(__line(iterator.next().value), "10");
__check(__line(iterator.next().value), "20");
__check(__line(iterator.next().value), "30");
