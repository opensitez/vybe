// vybe-test: js/structured_clone_advanced/structuredclone_array_with_holes
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

const arr = [1, , 3];
const clone = structuredClone(arr);
__check(__line(clone.length), "3");
__check(__line(clone[0], clone[2]), "1 3");
