// vybe-test: js/object_literal_advanced/shorthand_property
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

const x = 1, y = 2;
const point = { x, y };
__check(__line(point.x), "1");
__check(__line(point.y), "2");
