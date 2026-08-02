// vybe-test: js/functional_fp_patterns/lens_pattern
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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

const lens = (getter, setter) => ({ get: getter, set: setter });
const view = (l, obj) => l.get(obj);
const set = (l, val, obj) => l.set(val, obj);
const over = (l, fn, obj) => set(l, fn(view(l, obj)), obj);

const nameLens = lens(o => o.name, (v, o) => ({...o, name: v}));
const person = { name: "Alice", age: 30 };
__check(__line(view(nameLens, person)), "Alice");
const updated = over(nameLens, n => n.toUpperCase(), person);
__check(__line(updated.name), "ALICE");
__check(__line(person.name), "Alice");
