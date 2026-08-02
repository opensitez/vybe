// vybe-test: js/control_flow_patterns/for_of_with_index_via_entries
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

const arr = ["a", "b", "c"];
const result = [];
for (const [i, v] of arr.entries()) {
    result.push(i + ":" + v);
}
console.log(result.join(","));
