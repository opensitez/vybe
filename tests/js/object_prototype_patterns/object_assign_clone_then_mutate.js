// vybe-test: js/object_prototype_patterns/object_assign_clone_then_mutate
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

const original = { name: "Alice", score: 10 };
const copy = Object.assign({}, original);
copy.score += 5;
__check(__line(original.score), "10");
__check(__line(copy.score), "15");
