// vybe-test: js/typed_array_advanced/atomics_compareExchange
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const sab = new SharedArrayBuffer(4);
const ta = new Int32Array(sab);
ta[0] = 42;
const result = Atomics.compareExchange(ta, 0, 42, 99);
__check(__line(result), "42"); // old value
__check(__line(ta[0]), "99");  // new value — exchange happened
const result2 = Atomics.compareExchange(ta, 0, 42, 0); // expected wrong
__check(__line(result2), "99"); // old value (99)
__check(__line(ta[0]), "99");   // unchanged (99)
