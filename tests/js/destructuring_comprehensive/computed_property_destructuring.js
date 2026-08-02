// vybe-test: js/destructuring_comprehensive/computed_property_destructuring
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

const key = "name";
const { [key]: value } = { name: "Alice" };
__check(__line(value), "Alice");
const prop = "age";
const { [prop]: age = 25 } = {};
__check(__line(age), "25");
