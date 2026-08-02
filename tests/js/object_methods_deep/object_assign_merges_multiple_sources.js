// vybe-test: js/object_methods_deep/object_assign_merges_multiple_sources
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const target = { a: 1 };
const result = Object.assign(target, { b: 2 }, { c: 3 }, { b: 99 });
__check(__line(result === target), "true");
__check(__line(result.a + "," + result.b + "," + result.c), "1,99,3");
