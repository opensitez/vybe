// vybe-test: js/object_introspection/object_assign_copies_own_enumerable_only
// origin: languages/js/tests/js/test_object_introspection.rs

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

const src = { a: 1, b: 2 };
Object.defineProperty(src, "hidden", { value: 99, enumerable: false });
const dest = {};
Object.assign(dest, src);
__check(__line(dest.a), "1");
__check(__line(dest.b), "2");
__check(__line(dest.hidden), "undefined");
