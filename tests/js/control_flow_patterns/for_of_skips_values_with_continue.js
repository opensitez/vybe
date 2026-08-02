// vybe-test: js/control_flow_patterns/for_of_skips_values_with_continue
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

const seen = [];
for (const n of [1, 2, 3, 4]) {
    if (n % 2 === 0) continue;
    seen.push(n);
}
console.log(seen.join(","));
