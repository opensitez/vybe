// vybe-test: js/structured_clone_advanced/structuredclone_typed_array
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

const orig = new Uint8Array([1, 2, 3]);
const clone = structuredClone(orig);
clone[0] = 99;
__check(__line(orig[0]), "1");
__check(__line(clone[0]), "99");
