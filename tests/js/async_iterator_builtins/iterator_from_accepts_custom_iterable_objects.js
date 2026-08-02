// vybe-test: js/async_iterator_builtins/iterator_from_accepts_custom_iterable_objects
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

const source = {
  *[Symbol.iterator]() {
    yield "x";
    yield "y";
  }
};
const iterator = Iterator.from(source);
__check(__line(iterator.next().value + iterator.next().value), "xy");
