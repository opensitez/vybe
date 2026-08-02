// vybe-test: js/map_set_deep/map_preserves_insertion_order
// origin: languages/js/tests/js/test_map_set_deep.rs

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

const m = new Map();
m.set("c", 3); m.set("a", 1); m.set("b", 2);
const keys = [];
m.forEach((v, k) => keys.push(k));
console.log(keys.join(","));
