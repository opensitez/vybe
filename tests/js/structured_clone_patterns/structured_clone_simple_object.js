// vybe-test: js/structured_clone_patterns/structured_clone_simple_object
// origin: languages/js/tests/js/test_structured_clone_patterns.rs

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

const original = { a: 1, b: "hello", c: true };
const clone = structuredClone(original);
__check(__line(clone.a), "1");
__check(__line(clone.b), "hello");
__check(__line(clone.c), "true");
__check(__line(clone === original), "false");
