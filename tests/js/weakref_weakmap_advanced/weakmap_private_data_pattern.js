// vybe-test: js/weakref_weakmap_advanced/weakmap_private_data_pattern
// origin: languages/js/tests/js/test_weakref_weakmap_advanced.rs

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

const _private = new WeakMap();
class Counter {
  constructor() { _private.set(this, { count: 0 }); }
  increment() { _private.get(this).count++; }
  get value() { return _private.get(this).count; }
}
const c = new Counter();
c.increment();
c.increment();
c.increment();
__check(__line(c.value), "3");
