// vybe-test: js/control_flow_advanced/for_of_map_yields_key_value_pairs
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

const m = new Map([["a", 1], ["b", 2]]);
const pairs = [];
for (const [k, v] of m) pairs.push(k + "=" + v);
console.log(pairs.join(","));
