// vybe-test: js/array_algorithms/iterative_flat_deep_arrays
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function flatDeep(arr) {
    const stack = [...arr];
    const res = [];
    while (stack.length) {
        const next = stack.pop();
        if (Array.isArray(next)) stack.push(...next);
        else res.push(next);
    }
    return res.reverse();
}
console.log(flatDeep([1, [2, [3, [4]]]]).join(","));
