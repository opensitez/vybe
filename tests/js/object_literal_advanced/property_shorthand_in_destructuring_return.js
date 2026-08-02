// vybe-test: js/object_literal_advanced/property_shorthand_in_destructuring_return
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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

function getPoint() {
    const x = 3, y = 4;
    return { x, y };
}
const { x, y } = getPoint();
__check(__line(x), "3");
__check(__line(y), "4");
