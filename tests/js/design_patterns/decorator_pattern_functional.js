// vybe-test: js/design_patterns/decorator_pattern_functional
// origin: languages/js/tests/js/test_design_patterns.rs

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

function withLogging(fn, name) {
    return function(...args) {
        const result = fn(...args);
        __check(__line(`${name}(${args}) = ${result}`), "add(3,4) = 7");
        return result;
    };
}
const add = withLogging((a, b) => a + b, "add");
add(3, 4);
