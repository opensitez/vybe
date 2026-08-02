// vybe-test: js/scope_prototype/structured_clone_like
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let original = { a: 1, b: { c: 2 } };
let clone = JSON.parse(JSON.stringify(original));
clone.b.c = 99;
__check(__line(original.b.c), "2");
__check(__line(clone.b.c), "99");
