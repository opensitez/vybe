// vybe-test: js/weakmap_weakset_patterns/weakmap_stores_per_object_data
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const data = new WeakMap();
const obj1 = {};
const obj2 = {};
data.set(obj1, { id: 1 });
data.set(obj2, { id: 2 });
__check(__line(data.get(obj1).id), "1");
__check(__line(data.get(obj2).id), "2");
__check(__line(data.has(obj1)), "true");
