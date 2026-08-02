// vybe-test: js/destructuring_advanced/nested_destructure_function_return
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

function getUser() {
  return { id: 1, address: { city: "NYC", zip: "10001" } };
}
const { address: { city } } = getUser();
__check(__line(city), "NYC");
