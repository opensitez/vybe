// vybe-test: js/structured_clone_advanced/structuredclone_deep_copy
// origin: languages/js/tests/js/test_structured_clone_advanced.rs

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

const orig = { nested: { x: 1 } };
const clone = structuredClone(orig);
clone.nested.x = 99;
__check(__line(orig.nested.x), "1");
