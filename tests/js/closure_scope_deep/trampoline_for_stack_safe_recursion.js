// vybe-test: js/closure_scope_deep/trampoline_for_stack_safe_recursion
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function trampoline(fn) {
    return function(...args) {
        let result = fn(...args);
        while (typeof result === "function") result = result();
        return result;
    };
}

// Stack-safe sum via trampoline
function sum(n, acc = 0) {
    if (n === 0) return acc;
    return () => sum(n - 1, acc + n);
}

const safeSum = trampoline(sum);
console.log(safeSum(100));
