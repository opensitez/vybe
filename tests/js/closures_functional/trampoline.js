// vybe-test: js/closures_functional/trampoline
// origin: languages/js/tests/js/test_closures_functional.rs

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
        while (typeof result === "function") {
            result = result();
        }
        return result;
    };
}
function sumHelper(n, acc) {
    if (n === 0) return acc;
    return () => sumHelper(n - 1, acc + n);
}
let tSum = trampoline(sumHelper);
console.log(tSum(100, 0));
