// vybe-test: js/structured_clone_patterns/structured_clone_arraybuffer
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

const buffer = new ArrayBuffer(4);
const view = new Uint8Array(buffer);
view[0] = 42;
const clonedBuffer = structuredClone(buffer);
const clonedView = new Uint8Array(clonedBuffer);
__check(__line(clonedView[0]), "42");
clonedView[0] = 99;
__check(__line(view[0]), "42"); // original unchanged
