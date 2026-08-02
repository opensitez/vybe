// vybe-test: js/map_set_deep/map_iteration_via_for_of
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

const m = new Map([["x", 10], ["y", 20], ["z", 30]]);
const result = [];
for (const [k, v] of m) result.push(k + "=" + v);
console.log(result.join(","));
