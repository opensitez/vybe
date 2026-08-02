// vybe-test: js/structured_clone_patterns/structured_clone_map
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

const original = new Map([["key", { val: 1 }]]);
const clone = structuredClone(original);
clone.get("key").val = 99;
__check(__line(original.get("key").val), "1");
__check(__line(clone.get("key").val), "99");
