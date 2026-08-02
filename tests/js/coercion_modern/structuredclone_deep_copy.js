// vybe-test: js/coercion_modern/structuredclone_deep_copy
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let orig = { a: 1, b: { c: 2, d: [3, 4] } };
let clone = structuredClone(orig);
clone.b.c = 99;
clone.b.d.push(5);
__check(__line(orig.b.c), "2");
__check(__line(orig.b.d.length), "2");
__check(__line(clone.b.c), "99");
__check(__line(clone.b.d.length), "3");
