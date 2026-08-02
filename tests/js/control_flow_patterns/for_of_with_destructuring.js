// vybe-test: js/control_flow_patterns/for_of_with_destructuring
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

const pairs = [["a", 1], ["b", 2], ["c", 3]];
const result = [];
for (const [key, val] of pairs) {
    result.push(key + "=" + val);
}
console.log(result.join(","));
