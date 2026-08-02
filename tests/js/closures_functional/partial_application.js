// vybe-test: js/closures_functional/partial_application
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

function partial(fn, ...presets) {
    return function(...args) {
        return fn(...presets, ...args);
    };
}
function add3(a, b, c) { return a + b + c; }
let addTo10 = partial(add3, 3, 7);
__check(__line(addTo10(5)), "15");
__check(__line(addTo10(10)), "20");
