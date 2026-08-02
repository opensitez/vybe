// vybe-test: js/closures_functional/curry_string_formatter
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

function greet(greeting) {
    return function(name) {
        return greeting + ", " + name + "!";
    };
}
let hello = greet("Hello");
let hi = greet("Hi");
__check(__line(hello("Alice")), "Hello, Alice!");
__check(__line(hi("Bob")), "Hi, Bob!");
