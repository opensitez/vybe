// vybe-test: js/control_flow_patterns/for_of_array_iterator_manual
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

const it = [10, 20][Symbol.iterator]();
const res = [];
for (const x of it) res.push(x);
console.log(res.join(","));
