// vybe-test: js/control_flow_advanced/for_of_set_yields_values_in_insertion_order
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

const s = new Set([3, 1, 4, 1, 5]);
const vals = [];
for (const v of s) vals.push(v);
console.log(vals.join(","));
