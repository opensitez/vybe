// vybe-test: js/destructuring_patterns/deep_nested_mixed_destructuring
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

const data = {
    user: {
        name: "Bob",
        scores: [10, 20, 30]
    }
};
const { user: { name, scores: [first, , third] } } = data;
__check(__line(name), "Bob");
__check(__line(first), "10");
__check(__line(third), "30");
