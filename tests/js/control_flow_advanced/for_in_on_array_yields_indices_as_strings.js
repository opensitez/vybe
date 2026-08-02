// vybe-test: js/control_flow_advanced/for_in_on_array_yields_indices_as_strings
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

const arr = ["x", "y", "z"];
const indices = [];
for (const i in arr) {
    if (Object.prototype.hasOwnProperty.call(arr, i)) {
        indices.push(typeof i + ":" + i);
    }
}
console.log(indices.join("|"));
