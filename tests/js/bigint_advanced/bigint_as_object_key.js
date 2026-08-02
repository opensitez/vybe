// vybe-test: js/bigint_advanced/bigint_as_object_key
// origin: languages/js/tests/js/test_bigint_advanced.rs

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

const m = new Map();
m.set(1n, "one");
m.set(2n, "two");
__check(__line(m.get(1n)), "one");
__check(__line(m.size), "2");
