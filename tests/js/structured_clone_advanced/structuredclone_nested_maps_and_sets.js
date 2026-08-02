// vybe-test: js/structured_clone_advanced/structuredclone_nested_maps_and_sets
// origin: languages/js/tests/js/test_structured_clone_advanced.rs

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

const orig = { m: new Map([["k", new Set([1, 2])]]) };
const clone = structuredClone(orig);
clone.m.get("k").add(3);
__check(__line(orig.m.get("k").size), "2");
__check(__line(clone.m.get("k").size), "3");
