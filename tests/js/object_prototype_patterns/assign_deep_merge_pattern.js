// vybe-test: js/object_prototype_patterns/assign_deep_merge_pattern
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

// Note: Object.assign does shallow copy
const target = { a: { x: 1, y: 2 }, b: 10 };
const source = { a: { z: 3 }, b: 20 };
const merged = Object.assign({}, target, source);
// a is overwritten (not deep merged)
__check(__line(merged.b), "20");
__check(__line(merged.a.z), "3"); // has z from source
__check(__line(merged.a.x), "undefined"); // undefined — source.a replaced target.a
