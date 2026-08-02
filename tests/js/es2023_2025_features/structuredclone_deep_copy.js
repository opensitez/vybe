// vybe-test: js/es2023_2025_features/structuredclone_deep_copy
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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

const orig = { a: 1, b: { c: [1, 2, 3] } };
const copy = structuredClone(orig);
copy.b.c.push(4);
__check(__line(orig.b.c.length), "3");
__check(__line(copy.b.c.length), "4");
