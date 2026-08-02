// vybe-test: js/structured_clone_patterns/structured_clone_preserves_reference_graph
// origin: languages/js/tests/js/test_structured_clone_patterns.rs

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

const shared = { count: 0 };
const obj = { a: shared, b: shared };
const clone = structuredClone(obj);
// a and b should point to same cloned object
clone.a.count = 99;
__check(__line(clone.b.count), "99"); // 99 if same reference
__check(__line(obj.a.count), "0");   // 0 (not mutated)
