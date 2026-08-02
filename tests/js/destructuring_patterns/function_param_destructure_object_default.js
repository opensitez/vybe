// vybe-test: js/destructuring_patterns/function_param_destructure_object_default
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

function greet({ name = "World", greeting = "Hello" } = {}) {
    return `${greeting}, ${name}!`;
}
__check(__line(greet({ name: "Alice" })), "Hello, Alice!");
__check(__line(greet({ greeting: "Hi" })), "Hi, World!");
__check(__line(greet()), "Hello, World!");
