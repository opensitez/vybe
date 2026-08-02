// vybe-test: js/structured_clone_patterns/structured_clone_is_deep
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

const original = { nested: { x: 42 } };
const clone = structuredClone(original);
clone.nested.x = 99;
__check(__line(original.nested.x), "42");
__check(__line(clone.nested.x), "99");
