// vybe-test: js/json_serialization/json_deep_clone
// origin: languages/js/tests/js/test_json_serialization.rs

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

function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }
const orig = { a: { b: { c: [1, 2, 3] } } };
const clone = deepClone(orig);
clone.a.b.c.push(4);
__check(__line(orig.a.b.c.length), "3");
__check(__line(clone.a.b.c.length), "4");
__check(__line(orig.a === clone.a), "false");
