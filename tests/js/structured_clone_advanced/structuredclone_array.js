// vybe-test: js/structured_clone_advanced/structuredclone_array
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

const arr = [1, [2, 3], [4, [5]]];
const clone = structuredClone(arr);
clone[1][0] = 99;
__check(__line(arr[1][0]), "2");
__check(__line(clone[1][0]), "99");
