// vybe-test: js/destructuring_comprehensive/destructuring_rename_and_default
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

const { name: firstName = "Anonymous", age: years = 0 } = { name: "Alice", age: 30 };
__check(__line(firstName), "Alice");
__check(__line(years), "30");
const { x: a = 10, y: b = 20 } = { x: 5 };
__check(__line(a), "5");
__check(__line(b), "20");
